import pytest
import time
from datetime import datetime
import requests
from utils.voice_processor import VoiceProcessor, RateLimiter
from unittest.mock import patch, MagicMock


@pytest.fixture
def mock_config():
    with patch('utils.voice_processor.Config') as mock_config:
        mock_config.GEMINI_API_KEY = "test_key"
        mock_config.GEMINI_API_URL = "http://localhost/gemini"  # Použijeme localhost místo api.test.com
        mock_config.RATE_LIMIT_REQUESTS = 2
        mock_config.RATE_LIMIT_WINDOW = 1
        mock_config.GEMINI_CACHE_TTL = 3600
        yield mock_config


@pytest.fixture
def voice_processor(mock_config):
    with patch('requests_cache.CachedSession') as mock_cache:
        instance = mock_cache.return_value
        instance.post.return_value = MagicMock()
        processor = VoiceProcessor()
        processor.init_cache_session()  # Inicializujeme cache session
        yield processor


@pytest.fixture
def rate_limiter():
    return RateLimiter(max_requests=2, time_window=1)


@pytest.fixture
def mock_audio_file(tmp_path):
    audio_file = tmp_path / "test_audio.wav"
    audio_file.write_bytes(b"mock audio data")
    return str(audio_file)


def test_rate_limiter(rate_limiter):
    assert rate_limiter.can_make_request() is True
    rate_limiter.add_request()
    assert rate_limiter.can_make_request() is True
    rate_limiter.add_request()
    assert rate_limiter.can_make_request() is False
    time.sleep(1.1)
    assert rate_limiter.can_make_request() is True


@patch('requests.post')
def test_gemini_api_call_with_rate_limit(mock_post, voice_processor, mock_audio_file):
    mock_response = MagicMock()
    mock_response.json.return_value = {"text": "Test response"}
    mock_response.status_code = 200
    mock_post.return_value = mock_response

    result1 = voice_processor._call_gemini_api(mock_audio_file)
    assert "text" in result1
    assert result1["text"] == "Test response"
    
    result2 = voice_processor._call_gemini_api(mock_audio_file)
    assert "text" in result2

    # Simulujeme překročení rate limitu
    voice_processor.rate_limiter.requests.extend([time.time()] * 3)
    result3 = voice_processor._call_gemini_api(mock_audio_file)
    assert "error" in result3
    assert "Překročen rate limit" in result3["error"]


def test_gemini_api_caching(voice_processor, mock_audio_file):
    mock_response = MagicMock()
    mock_response.json.return_value = {"text": "Cached response"}
    mock_response.status_code = 200
    voice_processor.session.post.return_value = mock_response

    # První volání
    result1 = voice_processor._call_gemini_api(mock_audio_file)
    assert result1 == {"text": "Cached response"}
    
    # Druhé volání by mělo použít cache
    result2 = voice_processor._call_gemini_api(mock_audio_file)
    assert voice_processor.session.post.call_count == 1
    assert result1 == result2


def test_extract_entities(voice_processor):
    # Test pro částku a měnu
    text = "Přidej zálohu 1000 CZK pro Jana"
    with patch.object(voice_processor, '_load_employees', return_value=['Jan']):
        with patch.object(voice_processor.employee_manager, 'get_selected_employees', return_value=['Jan']):
            entities = voice_processor._extract_entities(text)
            assert entities["amount"] == 1000.0
            assert entities["currency"] == "CZK"
            assert entities["action"] == "add_advance"
            assert entities["employee"] == "Jan"

    # Test pro datum
    text = "Zaznamenej pracovní dobu pro 2025-05-06"
    entities = voice_processor._extract_entities(text)
    assert entities["date"] == "2025-05-06"
    assert entities["action"] == "record_time"

    # Test pro "dnes" a statistiky
    text = "Ukaž statistiky pro dnes"
    entities = voice_processor._extract_entities(text)
    assert entities["date"] == datetime.now().strftime("%Y-%m-%d")
    assert entities["action"] == "get_stats"


def test_normalize_date(voice_processor):
    assert voice_processor._normalize_date("2025-05-06") == "2025-05-06"
    assert voice_processor._normalize_date("06.05.2025") == "2025-05-06"
    assert voice_processor._normalize_date("06/05/2025") == "2025-05-06"
    assert voice_processor._normalize_date("invalid") is None


def test_validate_data(voice_processor):
    valid_data = {
        "employee": "Jan Novák",
        "date": "2025-05-06",
        "amount": 1000.0,
        "currency": "CZK",
        "action": "add_advance"
    }
    is_valid, errors = voice_processor._validate_data(valid_data)
    assert is_valid is True
    assert len(errors) == 0
    
    invalid_data = {
        "employee": None,
        "date": "invalid",
        "amount": -100,
        "currency": "USD",
        "action": None
    }
    is_valid, errors = voice_processor._validate_data(invalid_data)
    assert is_valid is False
    assert len(errors) > 0


def test_call_gemini_api_success(voice_processor, mock_audio_file):
    with patch('requests.Session.post') as mock_post:
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "text": "Zaznamenej pracovní dobu pro Jana Nováka na dnešek od 8:00 do 16:00",
            "confidence": 0.95
        }
        mock_post.return_value = mock_response

        result = voice_processor._call_gemini_api(mock_audio_file)
        assert "text" in result
        assert result["confidence"] == 0.95
        mock_post.assert_called_once()


def test_call_gemini_api_failure(voice_processor, mock_audio_file):
    with patch('requests.Session.post') as mock_post:
        mock_post.side_effect = requests.exceptions.RequestException("Internal Server Error")

        result = voice_processor._call_gemini_api(mock_audio_file)
        assert "error" in result
        assert "API volání selhalo" in result["error"]


def test_process_voice_command_success(voice_processor, mock_audio_file):
    mock_gemini_response = {
        "text": "Zaznamenej pracovní dobu pro Bláha Jakub dnes od 8:00 do 16:00",
        "confidence": 0.95
    }

    with patch.object(voice_processor, "_call_gemini_api", return_value=mock_gemini_response), \
         patch("utils.voice_processor.EmployeeManager.get_selected_employees", return_value=["Bláha Jakub"]), \
         patch.object(voice_processor, "_load_employees", return_value=["Bláha Jakub"]):
        
        voice_processor.employee_list = ["Bláha Jakub"]
        result = voice_processor.process_voice_command(mock_audio_file)

        assert result["success"] is True
        assert result["confidence"] == 0.95
        assert result["action"] == "record_time"
        assert result["employee"] == "Bláha Jakub"
        assert result["date"] == datetime.now().strftime("%Y-%m-%d")

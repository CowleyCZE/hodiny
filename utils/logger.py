import logging
from pathlib import Path

def setup_logger(name):
    """
    Vytvoří a nakonfiguruje logger s konzistentním formátem
    """
    logger = logging.getLogger(name)
    
    if not logger.handlers:
        logger.setLevel(logging.INFO)
        
        # Vytvoření log adresáře, pokud neexistuje
        log_dir = Path('logs')
        log_dir.mkdir(exist_ok=True)
        
        # File handler - použití Path pro spojování cest
        file_handler = logging.FileHandler(
            log_dir / f'{name}.log',
            encoding='utf-8'
        )
        file_handler.setLevel(logging.INFO)
        
        # Console handler
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        
        # Formát pro oba handlery
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)
        
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
    
    return logger

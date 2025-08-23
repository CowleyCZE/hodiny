/**
 * Theme Manager and Lazy Loading for hodiny application
 * Handles dark/light theme switching and implements lazy loading for better performance
 */

class ThemeManager {
    constructor() {
        this.currentTheme = this.getStoredTheme() || this.getSystemTheme();
        this.init();
    }

    init() {
        this.createThemeToggle();
        this.applyTheme(this.currentTheme);
        this.addEventListeners();
    }

    getSystemTheme() {
        return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
    }

    getStoredTheme() {
        return localStorage.getItem('hodiny-theme');
    }

    storeTheme(theme) {
        localStorage.setItem('hodiny-theme', theme);
    }

    applyTheme(theme) {
        document.documentElement.setAttribute('data-theme', theme);
        this.currentTheme = theme;
        this.storeTheme(theme);
        
        // Add transition class for smooth theme change
        document.body.classList.add('theme-transition');
        setTimeout(() => {
            document.body.classList.remove('theme-transition');
        }, 300);
    }

    toggleTheme() {
        const newTheme = this.currentTheme === 'light' ? 'dark' : 'light';
        this.applyTheme(newTheme);
        
        // Analytics event for theme switching
        if (typeof gtag !== 'undefined') {
            gtag('event', 'theme_switch', {
                'custom_parameter': newTheme
            });
        }
    }

    createThemeToggle() {
        const toggle = document.createElement('button');
        toggle.className = 'theme-toggle';
        toggle.setAttribute('aria-label', 'Přepnout téma');
        toggle.setAttribute('title', 'Přepnout mezi světlým a tmavým tématem');
        toggle.innerHTML = `
            <span class="sun-icon">☀️</span>
            <span class="moon-icon">🌙</span>
        `;
        
        document.body.appendChild(toggle);
        
        toggle.addEventListener('click', () => this.toggleTheme());
        
        // Keyboard support
        toggle.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                this.toggleTheme();
            }
        });
    }

    addEventListeners() {
        // Listen for system theme changes
        window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', (e) => {
            if (!this.getStoredTheme()) {
                this.applyTheme(e.matches ? 'dark' : 'light');
            }
        });
    }
}

class LazyLoader {
    constructor() {
        this.imageObserver = null;
        this.componentObserver = null;
        this.init();
    }

    init() {
        this.setupImageLazyLoading();
        this.setupComponentLazyLoading();
        this.setupFormLazyLoading();
    }

    setupImageLazyLoading() {
        if ('IntersectionObserver' in window) {
            this.imageObserver = new IntersectionObserver((entries) => {
                entries.forEach(entry => {
                    if (entry.isIntersecting) {
                        const img = entry.target;
                        if (img.dataset.src) {
                            img.src = img.dataset.src;
                            img.classList.add('loaded');
                            img.removeAttribute('data-src');
                            this.imageObserver.unobserve(img);
                        }
                    }
                });
            }, {
                rootMargin: '50px 0px'
            });

            // Observe all images with data-src
            document.querySelectorAll('img[data-src]').forEach(img => {
                img.classList.add('loading');
                this.imageObserver.observe(img);
            });
        } else {
            // Fallback for older browsers
            document.querySelectorAll('img[data-src]').forEach(img => {
                img.src = img.dataset.src;
                img.removeAttribute('data-src');
            });
        }
    }

    setupComponentLazyLoading() {
        if ('IntersectionObserver' in window) {
            this.componentObserver = new IntersectionObserver((entries) => {
                entries.forEach(entry => {
                    if (entry.isIntersecting) {
                        const component = entry.target;
                        this.loadComponent(component);
                        this.componentObserver.unobserve(component);
                    }
                });
            }, {
                rootMargin: '100px 0px'
            });

            // Observe components marked for lazy loading
            document.querySelectorAll('[data-lazy-component]').forEach(component => {
                this.componentObserver.observe(component);
            });
        }
    }

    setupFormLazyLoading() {
        // Lazy load heavy form components
        const heavyForms = document.querySelectorAll('.quick-entry-form, .excel-viewer-form');
        heavyForms.forEach(form => {
            this.optimizeFormElements(form);
        });
    }

    loadComponent(component) {
        const componentType = component.dataset.lazyComponent;
        
        switch (componentType) {
            case 'excel-viewer':
                this.loadExcelViewer(component);
                break;
            case 'chart':
                this.loadChart(component);
                break;
            case 'calendar':
                this.loadCalendar(component);
                break;
            default:
                component.classList.add('loaded');
        }
    }

    loadExcelViewer(component) {
        // Simulate loading Excel viewer component
        component.innerHTML = '<div class="loading">Načítám Excel viewer...</div>';
        
        setTimeout(() => {
            component.classList.add('loaded');
            component.innerHTML = component.dataset.content || '';
        }, 500);
    }

    loadChart(component) {
        // Lazy load chart libraries only when needed
        if (!window.Chart) {
            const script = document.createElement('script');
            script.src = 'https://cdn.jsdelivr.net/npm/chart.js';
            script.onload = () => {
                this.renderChart(component);
            };
            document.head.appendChild(script);
        } else {
            this.renderChart(component);
        }
    }

    renderChart(component) {
        // Chart rendering logic would go here
        component.classList.add('loaded');
    }

    loadCalendar(component) {
        // Lazy load calendar component
        component.classList.add('loaded');
    }

    optimizeFormElements(form) {
        // Debounce input events for better performance
        const inputs = form.querySelectorAll('input, select, textarea');
        inputs.forEach(input => {
            this.addDebounceEvent(input);
        });
    }

    addDebounceEvent(element) {
        let timeout;
        const originalHandler = element.oninput;
        
        element.oninput = function(event) {
            clearTimeout(timeout);
            timeout = setTimeout(() => {
                if (originalHandler) {
                    originalHandler.call(this, event);
                }
            }, 300);
        };
    }
}

class PerformanceOptimizer {
    constructor() {
        this.init();
    }

    init() {
        this.setupRequestIdleCallback();
        this.setupPreloadImportantResources();
        this.setupServiceWorkerCache();
        this.optimizeAnimations();
    }

    setupRequestIdleCallback() {
        if ('requestIdleCallback' in window) {
            requestIdleCallback(() => {
                this.performLowPriorityTasks();
            });
        } else {
            setTimeout(() => {
                this.performLowPriorityTasks();
            }, 2000);
        }
    }

    performLowPriorityTasks() {
        // Preload next likely navigation targets
        this.preloadNextPages();
        
        // Optimize images
        this.optimizeImages();
        
        // Cache frequently used data
        this.cacheFrequentData();
    }

    preloadNextPages() {
        const commonLinks = [
            '/zamestnanci',
            '/record_time',
            '/zalohy',
            '/excel_viewer'
        ];

        commonLinks.forEach(link => {
            const prefetchLink = document.createElement('link');
            prefetchLink.rel = 'prefetch';
            prefetchLink.href = link;
            document.head.appendChild(prefetchLink);
        });
    }

    optimizeImages() {
        // Convert images to WebP format when supported
        if (this.supportsWebP()) {
            document.querySelectorAll('img').forEach(img => {
                if (img.src && !img.src.includes('.webp')) {
                    const webpSrc = img.src.replace(/\.(jpg|jpeg|png)$/i, '.webp');
                    
                    // Check if WebP version exists
                    const testImg = new Image();
                    testImg.onload = () => {
                        img.src = webpSrc;
                    };
                    testImg.src = webpSrc;
                }
            });
        }
    }

    supportsWebP() {
        const canvas = document.createElement('canvas');
        canvas.width = 1;
        canvas.height = 1;
        return canvas.toDataURL('image/webp').indexOf('webp') > -1;
    }

    cacheFrequentData() {
        // Cache form data and settings
        const formData = {};
        document.querySelectorAll('form input, form select').forEach(input => {
            if (input.name && input.value) {
                formData[input.name] = input.value;
            }
        });
        
        if (Object.keys(formData).length > 0) {
            sessionStorage.setItem('hodiny-form-cache', JSON.stringify(formData));
        }
    }

    setupPreloadImportantResources() {
        // Preload critical CSS and JS
        const criticalResources = [
            '/static/css/style.css',
            '/static/js/confirmations.js'
        ];

        criticalResources.forEach(resource => {
            const link = document.createElement('link');
            link.rel = 'preload';
            link.as = resource.endsWith('.css') ? 'style' : 'script';
            link.href = resource;
            document.head.appendChild(link);
        });
    }

    setupServiceWorkerCache() {
        if ('serviceWorker' in navigator) {
            navigator.serviceWorker.register('/sw.js')
                .then(() => console.log('Service Worker registered'))
                .catch(() => console.log('Service Worker registration failed'));
        }
    }

    optimizeAnimations() {
        // Pause animations when tab is not visible
        document.addEventListener('visibilitychange', () => {
            const animations = document.getAnimations();
            if (document.hidden) {
                animations.forEach(animation => animation.pause());
            } else {
                animations.forEach(animation => animation.play());
            }
        });
    }
}

// Initialize modules when DOM is ready
document.addEventListener('DOMContentLoaded', () => {
    // Initialize theme manager
    window.themeManager = new ThemeManager();
    
    // Initialize lazy loader
    window.lazyLoader = new LazyLoader();
    
    // Initialize performance optimizer
    window.performanceOptimizer = new PerformanceOptimizer();
    
    // Add loading complete class
    document.body.classList.add('loaded');
    
    console.log('hodiny: Theme manager, lazy loader, and performance optimizer initialized');
});

// Export for module systems
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { ThemeManager, LazyLoader, PerformanceOptimizer };
}
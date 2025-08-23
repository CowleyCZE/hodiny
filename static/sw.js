/**
 * Service Worker for hodiny application
 * Provides offline functionality and performance improvements
 */

const CACHE_NAME = 'hodiny-v1.0.0';
const API_CACHE_NAME = 'hodiny-api-v1.0.0';

// Static assets to cache
const STATIC_ASSETS = [
    '/',
    '/static/style.css',
    '/static/css/responsive.css',
    '/static/css/themes.css',
    '/static/css/animations.css',
    '/static/js/modern-enhancements.js',
    '/static/js/confirmations.js',
    '/static/js/voice-handler.js',
    '/static/images/logo.png',
    '/templates/base.html'
];

// API endpoints to cache
const API_ENDPOINTS = [
    '/api/v1/health',
    '/api/v1/employees',
    '/api/v1/settings',
    '/api/v1/excel/status'
];

// Cache strategies
const CACHE_FIRST = 'cache-first';
const NETWORK_FIRST = 'network-first';
const NETWORK_ONLY = 'network-only';

// Route configurations
const ROUTE_CONFIG = {
    // Static assets - cache first
    '/static/': CACHE_FIRST,
    '/images/': CACHE_FIRST,
    
    // API endpoints - network first with cache fallback
    '/api/v1/health': NETWORK_FIRST,
    '/api/v1/employees': NETWORK_FIRST,
    '/api/v1/settings': NETWORK_FIRST,
    '/api/v1/excel/status': NETWORK_FIRST,
    
    // Dynamic endpoints - network only
    '/api/v1/time-entry': NETWORK_ONLY,
    '/upload': NETWORK_ONLY,
    '/send_email': NETWORK_ONLY
};

/**
 * Install event - cache static assets
 */
self.addEventListener('install', event => {
    console.log('Service Worker: Installing...');
    
    event.waitUntil(
        Promise.all([
            // Cache static assets
            caches.open(CACHE_NAME).then(cache => {
                console.log('Service Worker: Caching static assets');
                return cache.addAll(STATIC_ASSETS);
            }),
            
            // Cache API endpoints
            caches.open(API_CACHE_NAME).then(cache => {
                console.log('Service Worker: Pre-caching API endpoints');
                return Promise.all(
                    API_ENDPOINTS.map(url => 
                        fetch(url)
                            .then(response => {
                                if (response.ok) {
                                    return cache.put(url, response);
                                }
                            })
                            .catch(err => {
                                console.log(`Service Worker: Failed to cache ${url}:`, err);
                            })
                    )
                );
            })
        ]).then(() => {
            console.log('Service Worker: Installation complete');
            return self.skipWaiting();
        })
    );
});

/**
 * Activate event - clean up old caches
 */
self.addEventListener('activate', event => {
    console.log('Service Worker: Activating...');
    
    event.waitUntil(
        caches.keys().then(cacheNames => {
            return Promise.all(
                cacheNames.map(cacheName => {
                    if (cacheName !== CACHE_NAME && cacheName !== API_CACHE_NAME) {
                        console.log('Service Worker: Deleting old cache:', cacheName);
                        return caches.delete(cacheName);
                    }
                })
            );
        }).then(() => {
            console.log('Service Worker: Activation complete');
            return self.clients.claim();
        })
    );
});

/**
 * Fetch event - handle requests with appropriate cache strategy
 */
self.addEventListener('fetch', event => {
    const { request } = event;
    const url = new URL(request.url);
    
    // Skip non-GET requests for caching
    if (request.method !== 'GET') {
        return;
    }
    
    // Skip chrome-extension and other protocol requests
    if (!url.protocol.startsWith('http')) {
        return;
    }
    
    // Determine cache strategy
    const strategy = getCacheStrategy(url.pathname);
    
    event.respondWith(
        handleRequest(request, strategy)
    );
});

/**
 * Determine cache strategy for a given path
 */
function getCacheStrategy(pathname) {
    // Check exact matches first
    if (ROUTE_CONFIG[pathname]) {
        return ROUTE_CONFIG[pathname];
    }
    
    // Check prefix matches
    for (const [path, strategy] of Object.entries(ROUTE_CONFIG)) {
        if (pathname.startsWith(path)) {
            return strategy;
        }
    }
    
    // Default to network first for HTML pages
    if (pathname.endsWith('.html') || pathname === '/' || !pathname.includes('.')) {
        return NETWORK_FIRST;
    }
    
    // Default to cache first for static assets
    return CACHE_FIRST;
}

/**
 * Handle request based on cache strategy
 */
async function handleRequest(request, strategy) {
    const url = new URL(request.url);
    const cacheName = url.pathname.startsWith('/api/') ? API_CACHE_NAME : CACHE_NAME;
    
    switch (strategy) {
        case CACHE_FIRST:
            return handleCacheFirst(request, cacheName);
        
        case NETWORK_FIRST:
            return handleNetworkFirst(request, cacheName);
        
        case NETWORK_ONLY:
            return handleNetworkOnly(request);
        
        default:
            return handleNetworkFirst(request, cacheName);
    }
}

/**
 * Cache first strategy - try cache, fallback to network
 */
async function handleCacheFirst(request, cacheName) {
    try {
        const cache = await caches.open(cacheName);
        const cachedResponse = await cache.match(request);
        
        if (cachedResponse) {
            console.log('Service Worker: Serving from cache:', request.url);
            
            // Update cache in background
            fetch(request)
                .then(response => {
                    if (response.ok) {
                        cache.put(request, response.clone());
                    }
                })
                .catch(err => console.log('Service Worker: Background fetch failed:', err));
            
            return cachedResponse;
        }
        
        // Fallback to network
        console.log('Service Worker: Cache miss, fetching from network:', request.url);
        const networkResponse = await fetch(request);
        
        if (networkResponse.ok) {
            cache.put(request, networkResponse.clone());
        }
        
        return networkResponse;
        
    } catch (error) {
        console.error('Service Worker: Cache first error:', error);
        return createErrorResponse();
    }
}

/**
 * Network first strategy - try network, fallback to cache
 */
async function handleNetworkFirst(request, cacheName) {
    try {
        const cache = await caches.open(cacheName);
        
        try {
            console.log('Service Worker: Fetching from network:', request.url);
            const networkResponse = await fetch(request);
            
            if (networkResponse.ok) {
                cache.put(request, networkResponse.clone());
            }
            
            return networkResponse;
            
        } catch (networkError) {
            console.log('Service Worker: Network failed, trying cache:', request.url);
            const cachedResponse = await cache.match(request);
            
            if (cachedResponse) {
                console.log('Service Worker: Serving stale content from cache');
                return cachedResponse;
            }
            
            throw networkError;
        }
        
    } catch (error) {
        console.error('Service Worker: Network first error:', error);
        return createErrorResponse();
    }
}

/**
 * Network only strategy - always fetch from network
 */
async function handleNetworkOnly(request) {
    try {
        console.log('Service Worker: Network only:', request.url);
        return await fetch(request);
    } catch (error) {
        console.error('Service Worker: Network only error:', error);
        return createErrorResponse();
    }
}

/**
 * Create an error response for offline scenarios
 */
function createErrorResponse() {
    return new Response(
        JSON.stringify({
            error: 'Aplikace je momentálně offline',
            message: 'Zkontrolujte internetové připojení a zkuste to znovu',
            offline: true
        }),
        {
            status: 503,
            statusText: 'Service Unavailable',
            headers: {
                'Content-Type': 'application/json'
            }
        }
    );
}

/**
 * Handle background sync for offline actions
 */
self.addEventListener('sync', event => {
    console.log('Service Worker: Background sync triggered:', event.tag);
    
    if (event.tag === 'time-entry-sync') {
        event.waitUntil(syncTimeEntries());
    }
});

/**
 * Sync offline time entries when connection is restored
 */
async function syncTimeEntries() {
    try {
        // Get offline data from IndexedDB (would need to implement storage)
        const offlineEntries = await getOfflineTimeEntries();
        
        for (const entry of offlineEntries) {
            try {
                const response = await fetch('/api/v1/time-entry', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify(entry)
                });
                
                if (response.ok) {
                    await removeOfflineEntry(entry.id);
                    console.log('Service Worker: Synced offline entry:', entry.id);
                }
            } catch (error) {
                console.error('Service Worker: Failed to sync entry:', error);
            }
        }
    } catch (error) {
        console.error('Service Worker: Sync failed:', error);
    }
}

/**
 * Placeholder for getting offline entries (would implement with IndexedDB)
 */
async function getOfflineTimeEntries() {
    // This would be implemented with IndexedDB
    return [];
}

/**
 * Placeholder for removing synced offline entries
 */
async function removeOfflineEntry(entryId) {
    // This would be implemented with IndexedDB
    console.log('Service Worker: Would remove offline entry:', entryId);
}

/**
 * Handle push notifications (for future implementation)
 */
self.addEventListener('push', event => {
    console.log('Service Worker: Push notification received');
    
    const options = {
        body: 'Připomínka: Nezapomeňte zaznamenat pracovní dobu',
        icon: '/static/images/logo.png',
        badge: '/static/images/logo.png',
        tag: 'time-reminder',
        requireInteraction: true,
        actions: [
            {
                action: 'open',
                title: 'Otevřít aplikaci'
            },
            {
                action: 'dismiss',
                title: 'Zavřít'
            }
        ]
    };
    
    event.waitUntil(
        self.registration.showNotification('hodiny - Evidence pracovní doby', options)
    );
});

/**
 * Handle notification clicks
 */
self.addEventListener('notificationclick', event => {
    console.log('Service Worker: Notification clicked:', event.action);
    
    event.notification.close();
    
    if (event.action === 'open') {
        event.waitUntil(
            clients.openWindow('/')
        );
    }
});

console.log('Service Worker: Script loaded');
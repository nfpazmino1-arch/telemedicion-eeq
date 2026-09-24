const CACHE_NAME = "telemedicion-eeq-v4";

const ARCHIVOS = [
    "./",
    "./index.html",
    "./style.css",
    "./app.js",
    "./manifest.json",
    "./xlsx.full.min.js",
    "./icon-192.png",
    "./icon-512.png"
];


// INSTALACIÓN
self.addEventListener("install", event => {

    event.waitUntil(

        caches.open(CACHE_NAME)
            .then(cache => {

                return cache.addAll(ARCHIVOS);

            })

    );

    self.skipWaiting();

});


// ACTIVACIÓN
self.addEventListener("activate", event => {

    event.waitUntil(

        caches.keys()
            .then(keys => {

                return Promise.all(

                    keys
                        .filter(key =>
                            key !== CACHE_NAME
                        )
                        .map(key =>
                            caches.delete(key)
                        )

                );

            })

    );

    self.clients.claim();

});


// FUNCIONAMIENTO SIN INTERNET
self.addEventListener("fetch", event => {

    event.respondWith(

        caches.match(event.request)
            .then(respuesta => {

                return respuesta ||
                    fetch(event.request);

            })
            .catch(() => {

                return caches.match(
                    "./index.html"
                );

            })

    );

});
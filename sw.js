/* Network-first on purpose. A cached-stale timesheet is worse than a slow one:
 * serving old app code against new data has broken this kind of app before.
 * The cache exists only so the form still opens with no signal.
 *
 * Bump CACHE and the ?v= query strings in index.html together on any shell change.
 */
var CACHE = "wh-v8";
var SHELL = [
  "./index.html",
  "./styles.css?v=9",
  "./config.js?v=9",
  "./excel.js?v=9",
  "./app.js?v=9",
  "./manifest.json?v=9"
];

self.addEventListener("install", function (e) {
  e.waitUntil(caches.open(CACHE).then(function (c) {
    return c.addAll(SHELL);
  }).then(function () { return self.skipWaiting(); }));
});

self.addEventListener("activate", function (e) {
  e.waitUntil(caches.keys().then(function (keys) {
    return Promise.all(keys.map(function (k) {
      return k === CACHE ? null : caches.delete(k);
    }));
  }).then(function () { return self.clients.claim(); }));
});

self.addEventListener("fetch", function (e) {
  var url = new URL(e.request.url);

  // Never touch sign-in or Graph traffic.
  if (url.origin !== location.origin) return;
  if (e.request.method !== "GET") return;

  e.respondWith(
    fetch(e.request).then(function (res) {
      var copy = res.clone();
      caches.open(CACHE).then(function (c) { c.put(e.request, copy); });
      return res;
    }).catch(function () {
      return caches.match(e.request).then(function (hit) {
        return hit || caches.match("./index.html");
      });
    })
  );
});

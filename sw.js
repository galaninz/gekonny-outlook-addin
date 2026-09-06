/* Gekonny Subject Builder — service worker.
   Network-first for our own files, so edits reach everyone immediately.
   Cache is only a fallback for offline. Requests to Power Automate are
   never touched. */

var CACHE = "gekonny-sb-v2";
var SHELL = [
  "./taskpane.html",
  "./taskpane.css",
  "./taskpane.js",
  "./manifest.webmanifest",
  "./assets/icon-192.png",
  "./assets/icon-512.png"
];

self.addEventListener("install", function (e) {
  self.skipWaiting();
  e.waitUntil(
    caches.open(CACHE).then(function (c) { return c.addAll(SHELL); }).catch(function () {})
  );
});

self.addEventListener("activate", function (e) {
  e.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.map(function (k) { return k === CACHE ? null : caches.delete(k); }));
    }).then(function () { return self.clients.claim(); })
  );
});

self.addEventListener("fetch", function (e) {
  var req = e.request;
  if (req.method !== "GET") { return; }

  var url;
  try { url = new URL(req.url); } catch (err) { return; }
  if (url.origin !== self.location.origin) { return; }

  /* Network-first is not enough on its own: a plain fetch() may still be
     answered from the browser's own HTTP cache, and the host serves these
     files with a ten-minute max-age. That is how a published fix can sit
     on the server while the panel keeps running yesterday's code. Forcing
     revalidation makes the request conditional — a 304 when nothing
     changed, so the cost is one round trip, and never a stale script. */
  var fresh;
  try { fresh = new Request(req, { cache: "no-cache" }); } catch (err2) { fresh = req; }

  e.respondWith(
    fetch(fresh).then(function (res) {
      var copy = res.clone();
      caches.open(CACHE).then(function (c) { c.put(req, copy); }).catch(function () {});
      return res;
    }).catch(function () {
      return caches.match(req).then(function (hit) {
        return hit || caches.match("./taskpane.html");
      });
    })
  );
});

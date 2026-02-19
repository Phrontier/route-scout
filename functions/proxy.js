const ALLOWED_HOSTS = new Set([
  "raw.githubusercontent.com",
  "cdn.jsdelivr.net",
  "ourairports.com",
  "www.ourairports.com",
  "www.faa.gov",
  "faa.gov",
  "aeronav.faa.gov",
  "api.aviationapi.com",
  "www.aviationapi.com",
  "api.open-meteo.com",
  "archive-api.open-meteo.com"
]);

export async function onRequest(context) {
  const req = context.request;
  if (req.method === "OPTIONS") {
    return new Response(null, {
      status: 204,
      headers: {
        "access-control-allow-origin": "*",
        "access-control-allow-methods": "GET, OPTIONS",
        "access-control-allow-headers": "Content-Type",
        "cache-control": "public, max-age=300"
      }
    });
  }
  const reqUrl = new URL(req.url);
  const target = reqUrl.searchParams.get("url");

  if (!target) {
    return new Response("Missing url parameter", { status: 400 });
  }

  let parsed;
  try {
    parsed = new URL(target);
  } catch {
    return new Response("Invalid url parameter", { status: 400 });
  }

  if (!["http:", "https:"].includes(parsed.protocol)) {
    return new Response("Unsupported protocol", { status: 400 });
  }

  if (!ALLOWED_HOSTS.has(parsed.hostname)) {
    return new Response("Host not allowed", { status: 403 });
  }

  try {
    const upstream = await fetch(parsed.toString(), { method: "GET" });

    const headers = new Headers();
    const contentType = upstream.headers.get("content-type") || "text/plain; charset=utf-8";
    const cacheControl = upstream.headers.get("cache-control") || "public, max-age=300";

    headers.set("content-type", contentType);
    headers.set("cache-control", cacheControl);
    headers.set("access-control-allow-origin", "*");
    headers.set("access-control-allow-methods", "GET, OPTIONS");
    headers.set("access-control-allow-headers", "Content-Type");

    const body = await upstream.arrayBuffer();

    return new Response(body, {
      status: upstream.status,
      headers
    });
  } catch (error) {
    return new Response(`Proxy fetch failed: ${error?.message || "unknown error"}`, {
      status: 502,
      headers: {
        "content-type": "text/plain; charset=utf-8"
      }
    });
  }
}

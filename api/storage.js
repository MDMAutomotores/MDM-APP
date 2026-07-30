import { Redis } from "@upstash/redis";

// Vercel KV inyecta automáticamente estas variables de entorno cuando
// conectás el proyecto a un KV Store desde el dashboard de Vercel.
const redis = new Redis({
  url: process.env.KV_REST_API_URL,
  token: process.env.KV_REST_API_TOKEN,
});

const PREFIX = "mdm:";

export default async function handler(req, res) {
  const { key } = req.query;

  if (req.method === "GET") {
    if (!key) return res.status(400).json({ error: "falta key" });
    const value = await redis.get(PREFIX + key);
    if (value === null || value === undefined) {
      return res.status(404).json({ error: "not found" });
    }
    return res.status(200).json({ key, value, shared: true });
  }

  if (req.method === "POST") {
    const { key: k, value } = req.body || {};
    if (!k) return res.status(400).json({ error: "falta key" });
    await redis.set(PREFIX + k, value);
    return res.status(200).json({ key: k, value, shared: true });
  }

  if (req.method === "DELETE") {
    if (!key) return res.status(400).json({ error: "falta key" });
    await redis.del(PREFIX + key);
    return res.status(200).json({ key, deleted: true, shared: true });
  }

  return res.status(405).json({ error: "método no permitido" });
}

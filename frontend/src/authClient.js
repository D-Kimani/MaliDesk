const API_BASE = "https://malidesk.onrender.com";

export async function api(path, options = {}) {
  const url = `${API_BASE}${path}`;
  const response = await fetch(url, {
    credentials: "include",
    ...options,
    headers: {
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(data.error || `Request failed (${response.status})`);
  return data;
}
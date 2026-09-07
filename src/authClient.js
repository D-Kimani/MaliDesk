const API_BASE = (() => {
  const configured = String(import.meta.env.VITE_API_BASE_URL || "").trim().replace(/\/$/, "");
  if (configured) return configured;
  if (import.meta.env.PROD) return "https://malidesk.onrender.com";
  return "";
})();

export async function api(path, options = {}) {
  const response = await fetch(`${API_BASE}${path}`, {
    credentials: "include",
    ...options,
    headers: { "Content-Type": "application/json", ...(options.headers || {}) },
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(data.error || `Request failed (${response.status})`);
  return data;
}
export const login = (username, password, remember = false) => api("/api/auth/login", { method: "POST", body: JSON.stringify({ username, password, remember }) });
export const logout = () => api("/api/auth/logout", { method: "POST" });
export const me = () => api("/api/auth/me");

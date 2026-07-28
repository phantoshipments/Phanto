// src/pushSubscribe.js
import { supabase } from "./supabaseClient";

const VAPID_PUBLIC_KEY = import.meta.env.VITE_VAPID_PUBLIC_KEY;

function urlBase64ToUint8Array(base64String) {
  const padding = "=".repeat((4 - (base64String.length % 4)) % 4);
  const base64 = (base64String + padding).replace(/-/g, "+").replace(/_/g, "/");
  const raw = atob(base64);
  return Uint8Array.from([...raw].map((c) => c.charCodeAt(0)));
}

// Call this after Notification.requestPermission() resolves to "granted".
export async function subscribeToPush(coordinatorId) {
  if (!("serviceWorker" in navigator) || !("PushManager" in window)) return;
  const reg = await navigator.serviceWorker.register("/sw.js");
  const sub = await reg.pushManager.subscribe({
    userVisibleOnly: true,
    applicationServerKey: urlBase64ToUint8Array(VAPID_PUBLIC_KEY),
  });
  const json = sub.toJSON();
  await supabase.from("push_subscriptions").upsert(
    { coordinator_id: coordinatorId, endpoint: json.endpoint, p256dh: json.keys.p256dh, auth: json.keys.auth },
    { onConflict: "endpoint" }
  );
}

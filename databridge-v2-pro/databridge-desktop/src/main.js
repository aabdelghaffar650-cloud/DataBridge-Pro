const BACKEND_URL = "http://localhost:8504";
const POLL_INTERVAL_MS = 500;

async function isBackendReady() {
  try {
    const res = await fetch(BACKEND_URL, { mode: "no-cors" });
    return true;
  } catch {
    return false;
  }
}

async function waitForBackendAndLoad() {
  while (!(await isBackendReady())) {
    await new Promise((resolve) => setTimeout(resolve, POLL_INTERVAL_MS));
  }
  window.location.href = BACKEND_URL;
}

window.addEventListener("DOMContentLoaded", () => {
  waitForBackendAndLoad();
});

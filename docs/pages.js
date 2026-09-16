const appUrl = "https://383999cd38e747c5a16b6548816b4590.app-tencent.workbuddy.link";
const stayOnLanding = new URLSearchParams(window.location.search).get("landing") === "1";

if (!stayOnLanding) {
  window.location.replace(appUrl);
}

document.getElementById("year").textContent = new Date().getFullYear();

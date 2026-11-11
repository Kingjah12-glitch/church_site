const CONFIG = {
  address: "35 Onwe Road, Abakaliki, Ebonyi State, Nigeria",

  /* day: 0 = Sunday ... 6 = Saturday. h = hour (24h), m = minutes */
  services: [
    { day: 0, h: 8,  m: 45, endH: 12, endM: 0, name: "Sunday Service", note: "Songs, Bible study, sermon, the Lord’s Supper and more." },
    { day: 2, h: 16, m: 50, endH: 18, endM: 0, name: "Sisters’ Class", note: "A class for the sisters of the church." },
    { day: 3, h: 16, m: 50, endH: 18, endM: 0, name: "Bible Class",    note: "Study the scriptures together and ask your questions." },
    { day: 5, h: 16, m: 50, endH: 18, endM: 0, name: "Youth Class",    note: "Bible teaching and fellowship for the youth." }
  ],

  /* To add the preacher later, copy a line and fill in his name and number */
  contacts: [
    { title: "Youth President",    sub: "Enquiries about the youth class",    phone: "+2348146765418" },
    { title: "Sisters’ Enquiries", sub: "Enquiries about the sisters’ class", phone: "+2348032761204" }
  ]
};

const DAYS = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];
const $ = id => document.getElementById(id);

function fmtTime(h, m) {
  const ap = h >= 12 ? "PM" : "AM";
  return ((h + 11) % 12 + 1) + ":" + String(m).padStart(2, "0") + " " + ap;
}

/* Times page */
if ($("times")) {
  $("times").innerHTML = CONFIG.services.map(s =>
    `<div class="card gather">
       <h3>${s.name}</h3>
       <div class="d">${DAYS[s.day]}s, ${fmtTime(s.h, s.m)} to ${fmtTime(s.endH, s.endM)}</div>
       <p>${s.note}</p>
     </div>`).join("");
}

/* Contact page */
if ($("cards")) {
  $("cards").insertAdjacentHTML("beforeend", CONFIG.contacts.map(c => {
    const digits = c.phone.replace(/\D/g, "");
    return `<div class="card info">
      <h3>${c.title}</h3>
      <p>${c.sub}</p>
      <div class="links">
        <a href="tel:${c.phone}">Call ${c.phone}</a>
        <a href="https://wa.me/${digits}" target="_blank" rel="noopener">WhatsApp</a>
      </div>
    </div>`;
  }).join(""));
  $("mapLink").href = "https://www.google.com/maps/search/?api=1&query=" + encodeURIComponent(CONFIG.address);
}

/* Home page: countdown to the next gathering */
function nextService() {
  const now = new Date();
  let best = null;
  CONFIG.services.forEach(s => {
    const t = new Date(now);
    t.setHours(s.h, s.m, 0, 0);
    t.setDate(t.getDate() + ((s.day - now.getDay() + 7) % 7));
    if (t <= now) t.setDate(t.getDate() + 7);
    if (!best || t < best.t) best = { t, s };
  });
  return best;
}

if ($("cd")) {
  const tick = () => {
    const n = nextService();
    const diff = n.t - new Date();
    $("nextName").textContent = n.s.name;
    $("nextWhen").textContent = DAYS[n.s.day] + ", " + fmtTime(n.s.h, n.s.m) + " to " + fmtTime(n.s.endH, n.s.endM);
    $("cd").textContent = Math.floor(diff / 864e5);
    $("ch").textContent = Math.floor(diff % 864e5 / 36e5);
    $("cm").textContent = Math.floor(diff % 36e5 / 6e4);
    $("cs").textContent = Math.floor(diff % 6e4 / 1000);
  };
  tick();
  setInterval(tick, 1000);
}
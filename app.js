const VIDEOS = [
    { title: "Building a YouTube Clone with Tailwind CSS", channel: "DevChannel", views: "1.2M views", time: "2 days ago", duration: "10:00", category: "coding", seed: 1, avatar: "#6366f1" },
    { title: "Learn React in 10 Minutes", channel: "CodeMaster", views: "500K views", time: "5 hours ago", duration: "10:00", category: "coding", seed: 2, avatar: "#8b5cf6" },
    { title: "Top 10 VS Code Extensions", channel: "ToolTips", views: "80K views", time: "1 week ago", duration: "8:24", category: "coding", seed: 3, avatar: "#0ea5e9" },
    { title: "Modern Web Development 2026", channel: "WebWizard", views: "2.5M views", time: "1 month ago", duration: "24:15", category: "coding", seed: 4, avatar: "#10b981" },
    { title: "Beautiful CSS Layouts You Must Try", channel: "DesignPro", views: "300K views", time: "3 days ago", duration: "12:40", category: "coding", seed: 5, avatar: "#f59e0b" },
    { title: "Lo-fi Beats to Code To", channel: "ChillRadio", views: "4.8M views", time: "1 year ago", duration: "60:00", category: "music", seed: 6, avatar: "#ec4899" },
    { title: "Synthwave Mix 2026", channel: "NeonWaves", views: "890K views", time: "6 hours ago", duration: "45:30", category: "music", seed: 7, avatar: "#06b6d4" },
    { title: "Epic Gameplay - The Final Boss", channel: "ProGamer", views: "1.8M views", time: "2 weeks ago", duration: "18:52", category: "gaming", seed: 8, avatar: "#ef4444" },
    { title: "Speedrunning Records Compilation", channel: "FastRunner", views: "640K views", time: "4 days ago", duration: "15:07", category: "gaming", seed: 9, avatar: "#84cc16" },
    { title: "Live: Morning News Broadcast", channel: "XroNews", views: "120K watching", time: "live", duration: "LIVE", category: "live", seed: 10, avatar: "#f97316" },
    { title: "A Day in the Life of a Developer", channel: "LifeByCode", views: "950K views", time: "3 weeks ago", duration: "21:33", category: "vlog", seed: 11, avatar: "#3b82f6" },
    { title: "Travel Vlog: Hidden Gems of Vietnam", channel: "WanderLust", views: "2.1M views", time: "1 month ago", duration: "27:45", category: "vlog", seed: 12, avatar: "#14b8a6" },
];

const state = {
    query: "",
    filter: "all",
    modalOpen: false,
    playing: false,
};

const el = (id) => document.getElementById(id);
const $ = (sel) => document.querySelector(sel);

const reducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;

function thumbnail(v) {
    return `https://picsum.photos/seed/${v.seed}/320/180`;
}

function cardTemplate(v) {
    const isLive = v.duration === "LIVE";
    return `
        <div class="video-card cursor-pointer group" tabindex="0" role="link" data-id="${v.seed}" aria-label="${v.title} by ${v.channel}">
            <div class="relative mb-3 overflow-hidden rounded-xl">
                <img src="${thumbnail(v)}" alt="Thumbnail: ${v.title}" class="video-thumbnail w-full aspect-video object-cover rounded-xl group-hover:rounded-none transition-all duration-200" loading="lazy" decoding="async">
                <span class="thumbnail-overlay absolute inset-0 bg-black/30 opacity-0 transition-opacity duration-200 flex items-center justify-center">
                    <span class="material-icons-outlined text-white text-5xl drop-shadow-lg">play_circle</span>
                </span>
                <span class="absolute bottom-2 right-2 bg-black/80 text-white text-xs px-2 py-0.5 rounded font-medium tabular-nums ${isLive ? "bg-[#ff3333] font-semibold" : ""}">${v.duration}</span>
            </div>
            <div class="flex gap-3">
                <div class="w-9 h-9 rounded-full flex-shrink-0 flex items-center justify-center text-white text-xs font-bold" style="background:${v.avatar}">${v.channel[0]}</div>
                <div class="flex flex-col min-w-0">
                    <h3 class="font-semibold text-[15px] line-clamp-2 leading-snug">${v.title}</h3>
                    <p class="text-[#aaaaaa] text-[13px] mt-1 hover:text-white transition-colors duration-200 cursor-pointer">${v.channel}</p>
                    <p class="text-[#aaaaaa] text-[13px]">${v.views} &bull; ${v.time}</p>
                </div>
            </div>
        </div>`;
}

function renderVideos() {
    const grid = el("videoGrid");
    const query = state.query.trim().toLowerCase();
    const filtered = VIDEOS.filter(v => {
        const matchesFilter = state.filter === "all" || v.category === state.filter;
        const matchesQuery = !query || v.title.toLowerCase().includes(query) || v.channel.toLowerCase().includes(query);
        return matchesFilter && matchesQuery;
    });

    grid.innerHTML = filtered.map(cardTemplate).join("");
    el("emptyState").classList.toggle("hidden", filtered.length > 0);

    grid.querySelectorAll(".video-card").forEach(card => {
        card.addEventListener("click", () => openModal(Number(card.dataset.id)));
        card.addEventListener("keydown", (e) => {
            if (e.key === "Enter" || e.key === " ") {
                e.preventDefault();
                openModal(Number(card.dataset.id));
            }
        });
    });
}

function setActiveTag(tagEl) {
    document.querySelectorAll(".tag-chip").forEach(t => t.classList.toggle("active", t === tagEl));
}

function updateSidebarActive(item) {
    document.querySelectorAll(".sidebar-item").forEach(i => i.classList.toggle("active", i === item));
}

function showToast(message) {
    const toast = el("toast");
    toast.textContent = message;
    toast.classList.remove("hidden");
    if (!reducedMotion) {
        clearTimeout(showToast._t);
        showToast._t = setTimeout(() => toast.classList.add("hidden"), 2500);
    }
}

function openModal(id) {
    const v = VIDEOS.find(x => x.seed === id);
    if (!v) return;
    state.modalOpen = true;
    state.playing = false;
    el("modalTitle").textContent = v.title;
    el("modalChannel").textContent = v.channel;
    el("modalMeta").textContent = `${v.views} • ${v.time}`;
    el("modalAvatar").style.background = v.avatar;
    el("modalDescription").textContent = `Lorem ipsum dolor sit amet, consectetur adipiscing elit. ${v.title}. Join ${v.channel} and explore more great content on XroHub.`;
    el("playerThumb").src = thumbnail(v);
    el("playerThumb").alt = `Thumbnail: ${v.title}`;
    el("progressBar").style.width = "0%";
    el("playerTime").textContent = "0:00 / 10:00";
    state.lastVideoId = id;
    el("modal").classList.remove("hidden");
    document.body.style.overflow = "hidden";
    el("closeModalBtn").focus();
}

function closeModal() {
    if (!state.modalOpen) return;
    state.modalOpen = false;
    state.playing = false;
    clearInterval(togglePlay._timer);
    el("modal").classList.add("hidden");
    document.body.style.overflow = "";
    const card = document.querySelector(`.video-card[data-id="${state.lastVideoId}"]`);
    if (card) card.focus();
    else el("menuBtn").focus();
}

function togglePlay() {
    state.playing = !state.playing;
    const icon = el("playBtn").querySelector("span");
    icon.textContent = state.playing ? "pause" : "play_arrow";
    if (state.playing) {
        let progress = 0;
        clearInterval(togglePlay._timer);
        togglePlay._timer = setInterval(() => {
            progress = Math.min(progress + 2, 100);
            el("progressBar").style.width = progress + "%";
            if (progress >= 100) {
                clearInterval(togglePlay._timer);
                togglePlay();
                el("playBtn").querySelector("span").textContent = "replay";
            }
        }, 200);
    } else {
        clearInterval(togglePlay._timer);
    }
}

function setSidebarVisible(visible) {
    const sidebar = el("sidebar");
    const isDesktop = window.innerWidth >= 1024;
    if (isDesktop) {
        document.body.classList.toggle("sidebar-mini", !visible);
        const main = el("main");
        main.classList.toggle("lg:ml-60", visible);
        main.classList.toggle("lg:ml-[76px]", !visible);
    } else {
        sidebar.classList.toggle("-translate-x-full", !visible);
        sidebar.classList.toggle("translate-x-0", visible);
        el("backdrop").classList.toggle("hidden", !visible);
    }
}

document.addEventListener("DOMContentLoaded", () => {
    renderVideos();

    // Sidebar toggle
    let sidebarVisible = true;
    el("menuBtn").addEventListener("click", () => {
        if (window.innerWidth >= 1024) {
            sidebarVisible = !sidebarVisible;
            setSidebarVisible(sidebarVisible);
        } else {
            const sidebar = el("sidebar");
            const isOpen = !sidebar.classList.contains("-translate-x-full");
            if (isOpen) {
                el("backdrop").click();
            } else {
                sidebar.classList.remove("-translate-x-full");
                sidebar.classList.add("translate-x-0");
                el("backdrop").classList.remove("hidden");
            }
        }
    });

    el("backdrop").addEventListener("click", () => {
        const sidebar = el("sidebar");
        sidebar.classList.add("-translate-x-full");
        sidebar.classList.remove("translate-x-0");
        el("backdrop").classList.add("hidden");
    });

    // Tag chips
    document.querySelectorAll(".tag-chip").forEach(chip => {
        chip.addEventListener("click", () => {
            state.filter = chip.dataset.filter;
            setActiveTag(chip);
            renderVideos();
        });
        chip.addEventListener("keydown", (e) => {
            if (e.key === "Enter" || e.key === " ") {
                e.preventDefault();
                chip.click();
            }
        });
        chip.setAttribute("tabindex", "0");
        chip.setAttribute("role", "button");
    });

    // Sidebar nav
    document.querySelectorAll(".sidebar-item").forEach(item => {
        item.addEventListener("click", (e) => {
            e.preventDefault();
            updateSidebarActive(item);
            if (window.innerWidth < 1024) el("backdrop").click();
        });
    });

    // Search (desktop + mobile)
    el("searchForm").addEventListener("submit", (e) => {
        e.preventDefault();
        state.query = el("searchInput").value;
        renderVideos();
    });
    el("mobileSearchBtn").addEventListener("click", () => {
        el("mobileSearchBar").classList.toggle("hidden");
        el("mobileSearchInput").focus();
    });
    el("mobileSearchSubmit").addEventListener("click", () => {
        state.query = el("mobileSearchInput").value;
        renderVideos();
        el("mobileSearchBar").classList.add("hidden");
    });
    el("mobileSearchInput").addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
            state.query = e.target.value;
            renderVideos();
            el("mobileSearchBar").classList.add("hidden");
        }
    });

    // Demo buttons
    el("uploadBtn").addEventListener("click", () => showToast("Upload feature coming soon"));
    el("notifBtn").addEventListener("click", () => showToast("No new notifications"));
    document.querySelector('[aria-label="Voice search"]').addEventListener("click", () => showToast("Voice search coming soon"));

    // Modal
    el("closeModalBtn").addEventListener("click", closeModal);
    el("modalBackdrop").addEventListener("click", closeModal);
    el("playBtn").addEventListener("click", togglePlay);
    document.addEventListener("keydown", (e) => {
        if (e.key === "Escape") {
            if (state.modalOpen) closeModal();
            else {
                const sidebar = el("sidebar");
                if (window.innerWidth < 1024 && !sidebar.classList.contains("-translate-x-full")) el("backdrop").click();
            }
        }
    });
});

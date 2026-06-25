(() => {
    document.getElementById("app").innerHTML = `
        <svg class="loading-progress">
            <circle r="40%" cx="50%" cy="50%" />
            <circle r="40%" cx="50%" cy="50%" />
        </svg>
        <div class="loading-progress-text"></div>`;

    const scripts = [
        "_content/MudBlazor/MudBlazor.min.js",
        "download.js",
        "_framework/blazor.webassembly.js"
    ];

    const loadScript = src => new Promise((resolve, reject) => {
        const script = document.createElement("script");
        script.src = src;
        script.onload = resolve;
        script.onerror = () => reject(new Error(`Failed to load ${src}`));
        document.body.appendChild(script);
    });

    scripts.reduce((chain, src) => chain.then(() => loadScript(src)), Promise.resolve())
        .catch(error => {
            console.error(error);
            const errorUi = document.getElementById("blazor-error-ui");
            if (errorUi) {
                errorUi.style.display = "block";
            }
        });
})();

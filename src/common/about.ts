/**
 * Shows an About modal dialog with logo, version, and build date
 * @param moduleName - The name of the module (e.g., "Cash Flow Analysis")
 * @param version - The version number (e.g., "6.2.3")
 * @param versionDate - The build/version date (e.g., "02/03/2026")
 */
export function showAboutModal(moduleName: string, version: string, versionDate: string): void {
    // Create modal background
    const modalBg = document.createElement("div");
    modalBg.style.position = "fixed";
    modalBg.style.top = "0";
    modalBg.style.left = "0";
    modalBg.style.width = "100vw";
    modalBg.style.height = "100vh";
    modalBg.style.backgroundColor = "rgba(0,0,0,0.2)";
    modalBg.style.zIndex = "9000";
    modalBg.id = `about-modal-bg-${moduleName.replace(/\s+/g, "-").toLowerCase()}`;

    // Modal window
    const modal = document.createElement("div");
    modal.style.position = "absolute";
    modal.style.top = "50%";
    modal.style.left = "50%";
    modal.style.transform = "translate(-50%,-50%)";
    modal.style.background = "#fff";
    modal.style.boxShadow = "0 2px 12px rgba(0,0,0,0.25)";
    modal.style.borderRadius = "12px";
    modal.style.padding = "32px 24px 20px 24px";
    modal.style.textAlign = "center";
    modal.style.minWidth = "320px";
    modal.style.maxWidth = "95vw";
    modal.style.fontFamily = "'Segoe UI',sans-serif";
    modal.style.zIndex = "9100";
    modal.tabIndex = 0;

    // Logo
    const logo = document.createElement("img");
    logo.src = "../../../assets/agvantis_large.png";
    logo.alt = "AgAutomate Logo";
    logo.style.maxWidth = "160px";
    logo.style.marginBottom = "16px";
    logo.style.display = "block";
    logo.style.marginLeft = "auto";
    logo.style.marginRight = "auto";
    logo.onerror = function () { 
        logo.style.display = "none"; 
    };

    // Heading
    const heading = document.createElement("div");
    heading.innerText = `AgAutomate/${moduleName}`;
    heading.style.fontSize = "1.25em";
    heading.style.fontWeight = "bold";
    heading.style.marginBottom = "12px";

    // Version
    const versionElem = document.createElement("div");
    versionElem.innerHTML = `<b>Version:</b> ${version}`;
    versionElem.style.marginBottom = "4px";

    // Date
    const dateElem = document.createElement("div");
    dateElem.innerHTML = `<b>Build Date:</b> ${versionDate}`;
    dateElem.style.marginBottom = "18px";

    // Close button
    const closeBtn = document.createElement("button");
    closeBtn.textContent = "Close";
    closeBtn.style.marginTop = "8px";
    closeBtn.style.fontSize = "1em";
    closeBtn.style.padding = "5px 18px";
    closeBtn.style.borderRadius = "7px";
    closeBtn.style.background = "#0078d4";
    closeBtn.style.color = "#fff";
    closeBtn.style.border = "none";
    closeBtn.style.cursor = "pointer";
    closeBtn.onclick = function() {
        document.body.removeChild(modalBg);
    };

    // Allow closing by clicking background
    modalBg.onclick = (e) => {
        // Only close if clicked on the background, not on the modal itself
        if (e.target === modalBg) {
            document.body.removeChild(modalBg);
        }
    };

    // Allow closing with Escape key
    const handleKeyDown = (e: KeyboardEvent) => {
        if (e.key === "Escape") {
            document.body.removeChild(modalBg);
            document.removeEventListener("keydown", handleKeyDown);
        }
    };
    document.addEventListener("keydown", handleKeyDown);

    modal.appendChild(logo);
    modal.appendChild(heading);
    modal.appendChild(versionElem);
    modal.appendChild(dateElem);
    modal.appendChild(closeBtn);

    modalBg.appendChild(modal);
    document.body.appendChild(modalBg);

    // Focus for accessibility
    setTimeout(() => modal.focus(), 100);
}

// import { goToWorksheet, clearRows } from "../";
// import { showConfirmDialog } from "../common/showConfirmDialog";
import { loadCashFlowModule } from "../modules/cashflow/cashflow";
import { getUserSettings } from "../common/userSettings";
import { MODULE_REGISTRY, ModuleMetadata, getModuleById } from "../common/moduleRegistry";
import { showSettingsDialog } from "./settings";
import settingsHtml from "./settings.html?raw";

// Moved arrayBufferToBase64 to src/common/textLib.ts

// ---------------------------------------------------------
// Office Entry Point
// ---------------------------------------------------------
Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("AgAutomate task pane ready");

    // Toggle UI: hide sideload message, show real UI
    const sideload = document.getElementById("sideload-msg");
    const appBody = document.getElementById("app-body");

    if (sideload) sideload.style.display = "none";
    if (appBody) appBody.style.display = "block";

    // Inject settings dialog HTML into body
    const body = document.body;
    if (body) {
      const tempDiv = document.createElement("div");
      tempDiv.innerHTML = settingsHtml;
      while (tempDiv.firstChild) {
        body.appendChild(tempDiv.firstChild);
      }
    }

    // Initialize settings and render menu
    await initializeSettings();

    // Wire up settings button
    const settingsBtn = document.getElementById("settings-btn");
    if (settingsBtn) {
      settingsBtn.addEventListener("click", () => {
        showSettingsDialog();
      });
    }

    // Listen for settings changes
    window.addEventListener("settingsSaved", async () => {
      await renderMenu();
    });

    // Your navigation + buttons + health check
    setupNavigation();
    //setupCashFlowButtons();
    checkServiceHealth();

    document.getElementById("refresh-health-btn")!.addEventListener("click", () => {
      checkServiceHealth();
    });
  }
});

/**
 * Initialize settings and render the menu
 */
async function initializeSettings(): Promise<void> {
    try {
        await renderMenu();
    } catch (error) {
        console.error("Error initializing settings:", error);
        // Fallback: show empty menu or all modules
        await renderMenu();
    }
}

/**
 * Render the menu based on enabled modules from settings
 */
async function renderMenu(): Promise<void> {
    const mainMenu = document.getElementById("main-menu");
    if (!mainMenu) {
        console.error("Main menu container not found");
        return;
    }

    try {
        const settings = await getUserSettings();
        const enabledModules = settings.enabledModules || [];

        // Clear existing menu buttons
        mainMenu.innerHTML = "";

        // Render enabled modules
        if (enabledModules.length === 0) {
            // Show message if no modules enabled
            const message = document.createElement("div");
            message.className = "ms-font-m";
            message.style.padding = "20px";
            message.style.textAlign = "center";
            message.style.color = "#666";
            message.innerHTML = `
                <p>No modules are currently enabled.</p>
                <p style="margin-top: 10px; font-size: 12px;">Click the ⚙ button in the header to enable modules.</p>
            `;
            mainMenu.appendChild(message);
        } else {
            // Render buttons for enabled modules
            enabledModules.forEach(moduleId => {
                const module = getModuleById(moduleId);
                if (module) {
                    const button = createModuleButton(module);
                    mainMenu.appendChild(button);
                }
            });
        }

        // Re-setup navigation after rendering
        setupNavigation();
    } catch (error) {
        console.error("Error rendering menu:", error);
    }
}

/**
 * Create a menu button for a module
 */
function createModuleButton(module: ModuleMetadata): HTMLButtonElement {
    const button = document.createElement("button");
    button.className = "menu-btn";
    button.setAttribute("data-tool", module.id);

    if (module.iconPath) {
        const icon = document.createElement("img");
        icon.src = module.iconPath;
        icon.alt = "";
        icon.style.height = "20px";
        icon.style.verticalAlign = "middle";
        icon.style.marginRight = "8px";
        button.appendChild(icon);
    }

    button.appendChild(document.createTextNode(module.displayName));
    return button;
}

function setupNavigation() {
    document.querySelectorAll(".menu-btn").forEach(btn => {
        btn.addEventListener("click", async (e) => {
            const tool = (e.target as HTMLElement).closest(".menu-btn")?.getAttribute("data-tool");
            if (!tool) return;

            const module = getModuleById(tool);
            if (module) {
                await module.loadFunction();
            } else {
                console.error(`Module not found: ${tool}`);
            }

            // Hide menu, show back button
            document.getElementById("main-menu")!.classList.add("hidden");
            document.getElementById("back-btn")!.style.display = "block";
        });
    });

    document.getElementById("back-btn")!.addEventListener("click", () => {
        // Clear module content
        document.getElementById("module-container")!.innerHTML = "";

        // Show menu again
        document.getElementById("main-menu")!.classList.remove("hidden");

        // Hide back button
        document.getElementById("back-btn")!.style.display = "none";
    });
}

function formatToolName(tool: string): string {
    return tool
        .replace(/-/g, " ")
        .replace(/\b\w/g, c => c.toUpperCase());
}

/* ------------------------------
   Health Check
--------------------------------*/
async function checkServiceHealth() {
    const statusPanel = document.getElementById("status-panel")!;
    const statusText = document.getElementById("status-text")!;

    statusPanel.className = "status-panel"; // reset classes
    statusText.textContent = "Checking service health...";

    try {
        const response = await fetch("http://localhost:5000/health");

        if (!response.ok) {
            statusText.textContent = "Service is NOT running";
            statusPanel.classList.add("status-bad");
            return;
        }

        const data = await response.json();

        if (data.status === "Healthy") {
            statusText.textContent = "AgAutomate.Service is running";
            statusPanel.classList.add("status-ok");
        } else {
            statusText.textContent = "Service is NOT healthy";
            statusPanel.classList.add("status-bad");
        }

    } catch (err) {
        statusText.textContent = "Service is NOT running";
        statusPanel.classList.add("status-bad");
    }
}

function showStatusMessage(message: string, type: "ok" | "error") {
    const statusPanel = document.getElementById("status-panel")!;
    const statusText = document.getElementById("status-text")!;

    statusPanel.className = "status-panel"; // reset classes
    statusText.textContent = message;

    if (type === "ok") {
        statusPanel.classList.add("status-ok");
    } else {
        statusPanel.classList.add("status-bad");
    }
}




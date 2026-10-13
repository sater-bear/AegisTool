// Settings Dialog Management

import { getUserSettings, saveUserSettings, UserSettings } from "../common/userSettings";
import { MODULE_REGISTRY, ModuleMetadata } from "../common/moduleRegistry";

let currentSettings: UserSettings | null = null;
let settingsChanged = false;

/**
 * Initialize and show the settings dialog
 */
export async function showSettingsDialog(): Promise<void> {
    const dialog = document.getElementById("settings-dialog");
    if (!dialog) {
        console.error("Settings dialog not found");
        return;
    }

    // Load current settings
    try {
        currentSettings = await getUserSettings();
        settingsChanged = false;
        populateModulesList(currentSettings.enabledModules);
        dialog.classList.remove("hidden");
    } catch (error) {
        console.error("Error loading settings:", error);
        // Even if loading fails, show dialog with defaults
        currentSettings = { enabledModules: [], moduleSettings: {} };
        settingsChanged = false;
        populateModulesList([]);
        dialog.classList.remove("hidden");
        
        // Show warning message
        const message = document.createElement("div");
        message.className = "ms-font-s";
        message.style.padding = "10px";
        message.style.marginBottom = "15px";
        message.style.backgroundColor = "#fff4e5";
        message.style.border = "1px solid #ffaa44";
        message.style.borderRadius = "4px";
        message.textContent = "Warning: Settings storage is not available. Changes may not be saved.";
        const settingsBody = document.querySelector(".settings-body");
        if (settingsBody && !settingsBody.querySelector(".settings-warning")) {
            message.classList.add("settings-warning");
            settingsBody.insertBefore(message, settingsBody.firstChild);
        }
    }

    // Wire up event handlers
    setupSettingsHandlers();
}

/**
 * Hide the settings dialog
 */
export function hideSettingsDialog(): void {
    const dialog = document.getElementById("settings-dialog");
    if (dialog) {
        dialog.classList.add("hidden");
    }
    settingsChanged = false;
}

/**
 * Populate the modules list with checkboxes
 */
function populateModulesList(enabledModules: string[]): void {
    const modulesList = document.getElementById("modules-list");
    if (!modulesList) {
        console.error("Modules list container not found");
        return;
    }

    modulesList.innerHTML = "";

    MODULE_REGISTRY.forEach((module: ModuleMetadata) => {
        const moduleItem = document.createElement("div");
        moduleItem.className = "module-item";

        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.id = `module-${module.id}`;
        checkbox.value = module.id;
        checkbox.checked = enabledModules.includes(module.id);
        checkbox.addEventListener("change", () => {
            settingsChanged = true;
        });

        const label = document.createElement("label");
        label.htmlFor = `module-${module.id}`;
        label.className = "module-label";
        
        if (module.iconPath) {
            const icon = document.createElement("img");
            icon.src = module.iconPath;
            icon.alt = "";
            icon.style.height = "20px";
            icon.style.verticalAlign = "middle";
            icon.style.marginRight = "8px";
            label.appendChild(icon);
        }
        
        label.appendChild(document.createTextNode(module.displayName));

        moduleItem.appendChild(checkbox);
        moduleItem.appendChild(label);
        modulesList.appendChild(moduleItem);
    });
}

/**
 * Setup event handlers for the settings dialog
 */
function setupSettingsHandlers(): void {
    // Close button
    const closeBtn = document.getElementById("settings-close");
    if (closeBtn) {
        closeBtn.onclick = () => {
            hideSettingsDialog();
        };
    }

    // Save button
    const saveBtn = document.getElementById("settings-save");
    if (saveBtn) {
        saveBtn.onclick = async () => {
            await saveSettings();
        };
    }

    // Reset button
    const resetBtn = document.getElementById("settings-reset");
    if (resetBtn) {
        resetBtn.onclick = () => {
            // Uncheck all modules in the UI (don't save yet)
            populateModulesList([]);
            settingsChanged = true; // Mark as changed so user can save if they want
        };
    }
}

/**
 * Get selected modules from checkboxes
 */
function getSelectedModules(): string[] {
    const checkboxes = document.querySelectorAll<HTMLInputElement>("#modules-list input[type='checkbox']:checked");
    return Array.from(checkboxes).map(cb => cb.value);
}

/**
 * Save settings
 */
async function saveSettings(): Promise<void> {
    if (!currentSettings) {
        return;
    }

    try {
        const selectedModules = getSelectedModules();
        currentSettings.enabledModules = selectedModules;

        await saveUserSettings(currentSettings);
        settingsChanged = false;

        // Notify that settings were saved - the taskpane will reload the menu
        const event = new CustomEvent("settingsSaved", { detail: currentSettings });
        window.dispatchEvent(event);

        hideSettingsDialog();
        
        // Show success message
        showSaveMessage("Settings saved successfully!");
    } catch (error) {
        console.error("Error saving settings:", error);
        const errorMsg = error instanceof Error ? error.message : "Unknown error";
        if (errorMsg.includes("not available") || errorMsg.includes("roamingSettings")) {
            alert("Settings storage is not available. Your changes cannot be saved. This may be due to Office.js not being fully initialized or the add-in running in an unsupported context.");
        } else {
            alert(`Failed to save settings: ${errorMsg}`);
        }
    }
}


/**
 * Show a temporary save message
 */
function showSaveMessage(message: string): void {
    // You could enhance this with a toast notification
    // For now, we'll just log it
    console.log(message);
}

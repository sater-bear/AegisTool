// User Settings Management
// Uses localStorage for user-level settings (works in Excel add-ins)
// Falls back to Office.js roamingSettings if available (for Outlook add-ins)

export interface UserSettings {
    // Module visibility - array of enabled module IDs
    enabledModules: string[];
    
    // Module-specific settings (extensible)
    // Format: { [moduleId]: { [settingKey]: value } }
    moduleSettings?: { [moduleId: string]: { [key: string]: any } };
}

const SETTINGS_KEY = "AgAutomate.UserSettings";
const SETTINGS_VERSION = 1;

// Default settings (no modules enabled by default)
const DEFAULT_SETTINGS: UserSettings = {
    enabledModules: [],
    moduleSettings: {}
};

/**
 * Check if localStorage is available
 */
function isLocalStorageAvailable(): boolean {
    try {
        const test = "__localStorage_test__";
        localStorage.setItem(test, test);
        localStorage.removeItem(test);
        return true;
    } catch {
        return false;
    }
}

/**
 * Check if roamingSettings is available (for Outlook add-ins)
 */
function isRoamingSettingsAvailable(): boolean {
    try {
        return !!(Office && Office.context && Office.context.roamingSettings);
    } catch {
        return false;
    }
}

/**
 * Get user settings from storage (localStorage or roamingSettings)
 * Returns default settings if none exist or if storage is unavailable
 */
export async function getUserSettings(): Promise<UserSettings> {
    return new Promise((resolve) => {
        try {
            // Try localStorage first (works in Excel add-ins)
            if (isLocalStorageAvailable()) {
                try {
                    const settingsJson = localStorage.getItem(SETTINGS_KEY);
                    
                    if (settingsJson) {
                        try {
                            const settings: UserSettings = JSON.parse(settingsJson);
                            
                            // Validate and merge with defaults
                            const validatedSettings: UserSettings = {
                                enabledModules: Array.isArray(settings.enabledModules) 
                                    ? settings.enabledModules 
                                    : DEFAULT_SETTINGS.enabledModules,
                                moduleSettings: settings.moduleSettings || DEFAULT_SETTINGS.moduleSettings
                            };
                            
                            resolve(validatedSettings);
                            return;
                        } catch (parseError) {
                            console.error("Error parsing settings from localStorage:", parseError);
                        }
                    }
                } catch (error) {
                    console.error("Error reading from localStorage:", error);
                }
            }

            // Fallback to roamingSettings if available (for Outlook)
            if (isRoamingSettingsAvailable()) {
                try {
                    const roamingSettings = Office.context.roamingSettings;
                    if (roamingSettings) {
                        const settingsJson = roamingSettings.get(SETTINGS_KEY);
                        
                        if (settingsJson) {
                            try {
                                const settings: UserSettings = JSON.parse(settingsJson as string);
                                
                                // Validate and merge with defaults
                                const validatedSettings: UserSettings = {
                                    enabledModules: Array.isArray(settings.enabledModules) 
                                        ? settings.enabledModules 
                                        : DEFAULT_SETTINGS.enabledModules,
                                    moduleSettings: settings.moduleSettings || DEFAULT_SETTINGS.moduleSettings
                                };
                                
                                resolve(validatedSettings);
                                return;
                            } catch (parseError) {
                                console.error("Error parsing settings from roamingSettings:", parseError);
                            }
                        }
                    }
                } catch (error) {
                    console.error("Error reading from roamingSettings:", error);
                }
            }
            
            // No settings found or storage unavailable - return defaults
            resolve({ ...DEFAULT_SETTINGS });
        } catch (error) {
            console.error("Error getting user settings:", error);
            // Return defaults instead of rejecting to prevent app breakage
            resolve({ ...DEFAULT_SETTINGS });
        }
    });
}

/**
 * Save user settings to storage (localStorage or roamingSettings)
 */
export async function saveUserSettings(settings: UserSettings): Promise<void> {
    return new Promise((resolve, reject) => {
        try {
            // Validate settings
            const validatedSettings: UserSettings = {
                enabledModules: Array.isArray(settings.enabledModules) 
                    ? settings.enabledModules 
                    : DEFAULT_SETTINGS.enabledModules,
                moduleSettings: settings.moduleSettings || DEFAULT_SETTINGS.moduleSettings
            };
            
            const settingsJson = JSON.stringify(validatedSettings);
            
            // Try localStorage first (works in Excel add-ins)
            if (isLocalStorageAvailable()) {
                try {
                    localStorage.setItem(SETTINGS_KEY, settingsJson);
                    resolve();
                    return;
                } catch (error) {
                    console.error("Error saving to localStorage:", error);
                    // Continue to try roamingSettings as fallback
                }
            }

            // Fallback to roamingSettings if available (for Outlook)
            if (isRoamingSettingsAvailable()) {
                try {
                    const roamingSettings = Office.context.roamingSettings;
                    if (roamingSettings) {
                        roamingSettings.set(SETTINGS_KEY, settingsJson);
                        
                        // Save to Office
                        roamingSettings.saveAsync((result) => {
                            if (result.status === Office.AsyncResultStatus.Succeeded) {
                                resolve();
                            } else {
                                reject(new Error(`Failed to save settings: ${result.error?.message || "Unknown error"}`));
                            }
                        });
                        return;
                    }
                } catch (error) {
                    console.error("Error saving to roamingSettings:", error);
                }
            }
            
            // No storage available
            reject(new Error("No storage mechanism available (localStorage or roamingSettings)"));
        } catch (error) {
            console.error("Error saving user settings:", error);
            reject(error);
        }
    });
}

/**
 * Get module-specific settings for a given module
 */
export async function getModuleSettings(moduleId: string): Promise<{ [key: string]: any }> {
    const settings = await getUserSettings();
    return settings.moduleSettings?.[moduleId] || {};
}

/**
 * Save module-specific settings for a given module
 */
export async function saveModuleSettings(moduleId: string, moduleSettings: { [key: string]: any }): Promise<void> {
    const settings = await getUserSettings();
    
    if (!settings.moduleSettings) {
        settings.moduleSettings = {};
    }
    
    settings.moduleSettings[moduleId] = moduleSettings;
    await saveUserSettings(settings);
}

/**
 * Reset settings to defaults
 */
export async function resetSettingsToDefaults(): Promise<void> {
    await saveUserSettings({ ...DEFAULT_SETTINGS });
}

// Module Registry - Centralized module metadata and configuration

import { loadCashFlowModule } from "../modules/cashflow/cashflow";
import { loadFCCGBillingsModule } from "../modules/fccg-billings/fccg-billings";

export interface ModuleMetadata {
    id: string;                    // matches data-tool attribute in HTML
    displayName: string;            // Display name for UI
    iconPath?: string;             // Path to icon (relative to assets)
    loadFunction: () => void | Promise<void>;      // Function to load the module (can be async)
}

// Registry of all available modules
export const MODULE_REGISTRY: ModuleMetadata[] = [
    {
        id: "cashflow",
        displayName: "Cash Flow Analysis",
        iconPath: "../../assets/a_leaf_32_white.png",
        loadFunction: loadCashFlowModule
    },
    {
        id: "chron",
        displayName: "Chron Report",
        iconPath: "../../assets/a_leaf_32_white.png",
        loadFunction: () => {
            // Placeholder - implement when module is created
            console.log("Chron module not yet implemented");
        }
    },
    {
        id: "participation",
        displayName: "Participation Exception Report",
        iconPath: "../../assets/a_leaf_32_white.png",
        loadFunction: () => {
            console.log("Participation module not yet implemented");
        }
    },
    {
        id: "fccg-billings",
        displayName: "FCCG - Billings & Fees",
        iconPath: "../../assets/a_leaf_32_white.png",
        loadFunction: loadFCCGBillingsModule
    },
    {
        id: "fccg-variance",
        displayName: "FCCG - Billing Variance",
        iconPath: "../../assets/a_leaf_32_white.png",
        loadFunction: () => {
            console.log("FCCG Variance module not yet implemented");
        }
    },
    {
        id: "file-maintenance",
        displayName: "File Maintenance Report",
        iconPath: "../../assets/a_leaf_32_white.png",
        loadFunction: () => {
            console.log("File Maintenance module not yet implemented");
        }
    },
    {
        id: "pdf",
        displayName: "PDF Tool",
        iconPath: "../../assets/a_leaf_32_white.png",
        loadFunction: () => {
            console.log("PDF module not yet implemented");
        }
    },
    {
        id: "vendor",
        displayName: "Vendor Tool",
        iconPath: "../../assets/a_leaf_32_white.png",
        loadFunction: () => {
            console.log("Vendor module not yet implemented");
        }
    },
    {
        id: "wire",
        displayName: "Wire Tool",
        iconPath: "../../assets/a_leaf_32_white.png",
        loadFunction: () => {
            console.log("Wire module not yet implemented");
        }
    },
    {
        id: "other",
        displayName: "Other",
        iconPath: "../../assets/a_leaf_32_white.png",
        loadFunction: () => {
            console.log("Other module not yet implemented");
        }
    }
];

/**
 * Get module metadata by ID
 */
export function getModuleById(id: string): ModuleMetadata | undefined {
    return MODULE_REGISTRY.find(module => module.id === id);
}

/**
 * Get all module IDs
 */
export function getAllModuleIds(): string[] {
    return MODULE_REGISTRY.map(module => module.id);
}

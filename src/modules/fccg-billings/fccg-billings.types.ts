// src/modules/fccg-billings/fccg-billings.types.ts

/**
 * Company information structure
 * Equivalent to VBA ClassCompany
 */
export interface Company {
    companyName: string;
    companyID: string;
    companyType: string;
    glPayableAccount: string;
    wireScope: string;
    glAccount_FCCG: string;
    accountingUnit_FCCG_FLCA: string;
    accountingUnit_FCCG_PCA: string;
    glAccount_IP: string;
    accountingUnit_IP_FLCA: string;
    accountingUnit_IP_PCA: string;
    includeJV: string | boolean;
    worksheetNameBillings: string;
    worksheetNameWireBilling: string;
    worksheetNameFees: string;
    worksheetNameWireFee: string;
    emailAddressFees_FCCG: string;
    emailBodyFee_FCCG: string;
    emailAddressBilling_FCCG: string;
    emailBodyBilling: string;
    emailAddressWire: string;
    emailBodyWireFee: string;
}

/**
 * Company map - keyed by CompanyName
 */
export type CompanyMap = Map<string, Company>;

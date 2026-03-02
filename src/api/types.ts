export interface GraphUser {
  id: string;
  displayName: string;
  userPrincipalName: string;
  mail: string | null;
  jobTitle: string | null;
  department: string | null;
  officeLocation: string | null;
  accountEnabled: boolean;
  createdDateTime: string | null;
  signInActivity?: {
    lastSignInDateTime: string | null;
    lastNonInteractiveSignInDateTime: string | null;
  };
}

export interface GraphGroup {
  id: string;
  displayName: string;
  description: string | null;
  mailEnabled: boolean;
  mailNickname: string;
  securityEnabled: boolean;
  groupTypes: string[];
  membershipRule: string | null;
  createdDateTime: string;
}

export interface GraphDirectoryObject {
  "@odata.type": string;
  id: string;
  displayName: string;
  userPrincipalName?: string;
}

export interface GraphSubscribedSku {
  id: string;
  skuId: string;
  skuPartNumber: string;
  appliesTo: string;
  capabilityStatus: string;
  consumedUnits: number;
  prepaidUnits: {
    enabled: number;
    suspended: number;
    warning: number;
  };
}

export interface GraphLicenseDetail {
  id: string;
  skuId: string;
  skuPartNumber: string;
  servicePlans: Array<{
    servicePlanId: string;
    servicePlanName: string;
    provisioningStatus: string;
  }>;
}

export interface ODataResponse<T> {
  value: T[];
  "@odata.nextLink"?: string;
  "@odata.count"?: number;
}

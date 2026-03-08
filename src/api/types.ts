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

export interface GraphMailMessage {
  id: string;
  subject: string | null;
  bodyPreview: string;
  from: {
    emailAddress: { name: string; address: string };
  } | null;
  toRecipients: Array<{
    emailAddress: { name: string; address: string };
  }>;
  receivedDateTime: string;
  isRead: boolean;
  hasAttachments: boolean;
  importance: string;
}

export interface GraphMailMessageFull extends GraphMailMessage {
  body: {
    contentType: string;
    content: string;
  };
  ccRecipients: Array<{
    emailAddress: { name: string; address: string };
  }>;
}

export interface GraphAttachment {
  id: string;
  name: string;
  contentType: string;
  size: number;
  isInline: boolean;
}

export interface GraphDriveItem {
  id: string;
  name: string;
  size: number;
  lastModifiedDateTime: string;
  webUrl: string;
  folder?: { childCount: number };
  file?: { mimeType: string };
  "@microsoft.graph.downloadUrl"?: string;
}

export interface GraphCalendarEvent {
  id: string;
  subject: string | null;
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
  organizer: {
    emailAddress: { name: string; address: string };
  } | null;
  attendees: Array<{
    emailAddress: { name: string; address: string };
    status: { response: string };
    type: string;
  }>;
  isOnlineMeeting: boolean;
  onlineMeetingUrl: string | null;
  onlineMeeting: {
    joinUrl: string;
  } | null;
  bodyPreview: string;
}

export interface GraphOnlineMeeting {
  id: string;
  subject: string | null;
  startDateTime: string;
  endDateTime: string;
  joinWebUrl: string;
  chatInfo?: { threadId: string; messageId: string };
  participants?: {
    organizer?: { identity: { user?: { id: string; displayName: string } } };
    attendees?: Array<{ identity: { user?: { id: string; displayName: string } } }>;
  };
}

export interface GraphTranscript {
  id: string;
  meetingId: string;
  meetingOrganizerId: string;
  createdDateTime: string;
  transcriptContentUrl: string;
}

export interface ODataResponse<T> {
  value: T[];
  "@odata.nextLink"?: string;
  "@odata.count"?: number;
}

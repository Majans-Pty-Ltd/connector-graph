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

export interface GraphFileAttachment extends GraphAttachment {
  "@odata.type": string;
  contentBytes: string;  // base64-encoded content
}

export interface GraphMailFolder {
  id: string;
  displayName: string;
  parentFolderId: string;
  childFolderCount: number;
  unreadItemCount: number;
  totalItemCount: number;
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

export interface GraphSendMailRecipient {
  emailAddress: { name?: string; address: string };
}

export interface GraphSendMailRequest {
  message: {
    subject: string;
    body: { contentType: "Text" | "HTML"; content: string };
    toRecipients: GraphSendMailRecipient[];
    ccRecipients?: GraphSendMailRecipient[];
    importance?: "low" | "normal" | "high";
  };
  saveToSentItems?: boolean;
}

export interface ODataResponse<T> {
  value: T[];
  "@odata.nextLink"?: string;
  "@odata.count"?: number;
}

// ── SharePoint ──

export interface GraphSite {
  id: string;
  displayName: string;
  name: string;
  webUrl: string;
  description: string | null;
  createdDateTime: string;
  lastModifiedDateTime: string;
  root?: Record<string, unknown>;
  siteCollection?: { hostname: string };
}

export interface GraphDrive {
  id: string;
  name: string;
  driveType: string;
  webUrl: string;
  description: string | null;
  createdDateTime: string;
  lastModifiedDateTime: string;
  quota?: {
    total: number;
    used: number;
    remaining: number;
    state: string;
  };
}

export interface GraphSearchResult {
  id: string;
  name: string;
  size: number;
  webUrl: string;
  lastModifiedDateTime: string;
  file?: { mimeType: string };
  folder?: { childCount: number };
  parentReference?: {
    driveId: string;
    path: string;
  };
}

// ── Planner ──

export interface GraphPlannerPlan {
  id: string;
  title: string;
  owner: string;
  createdDateTime: string;
  createdBy: {
    user: { id: string; displayName?: string };
  };
}

export interface GraphPlannerBucket {
  id: string;
  name: string;
  planId: string;
  orderHint: string;
}

export interface GraphPlannerTask {
  "@odata.etag"?: string;
  id: string;
  title: string;
  planId: string;
  bucketId: string | null;
  percentComplete: number;
  startDateTime: string | null;
  dueDateTime: string | null;
  createdDateTime: string;
  completedDateTime: string | null;
  priority: number;
  assignments: Record<string, { orderHint: string }>;
  appliedCategories: Record<string, boolean>;
  orderHint: string;
  createdBy: {
    user: { id: string; displayName?: string };
  };
}

// ── To Do ──

export interface GraphTodoList {
  id: string;
  displayName: string;
  isOwner: boolean;
  isShared: boolean;
  wellknownListName: string;
}

export interface GraphTodoTask {
  id: string;
  title: string;
  status: "notStarted" | "inProgress" | "completed" | "waitingOnOthers" | "deferred";
  importance: "low" | "normal" | "high";
  isReminderOn: boolean;
  createdDateTime: string;
  lastModifiedDateTime: string;
  completedDateTime?: { dateTime: string; timeZone: string } | null;
  dueDateTime?: { dateTime: string; timeZone: string } | null;
  body?: { content: string; contentType: string };
}

// ── Teams Chat ──

export interface GraphChat {
  id: string;
  topic: string | null;
  chatType: "oneOnOne" | "group" | "meeting";
  createdDateTime: string;
  lastUpdatedDateTime: string;
  webUrl: string | null;
}

export interface GraphChatMessage {
  id: string;
  createdDateTime: string;
  lastModifiedDateTime: string;
  messageType: string;
  from: {
    user?: { id: string; displayName: string };
    application?: { id: string; displayName: string };
  } | null;
  body: {
    contentType: string;
    content: string;
  };
  importance: string;
  attachments: Array<{
    id: string;
    contentType: string;
    name: string | null;
  }>;
}

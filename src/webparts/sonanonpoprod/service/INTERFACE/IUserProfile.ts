// src/service/INTERFACE/IUserProfile.ts

/**
 * SharePoint User Profile property item
 * Example: { Key: "SPS-Location", Value: "Mumbai" }
 */
export interface IUserProfileProperty {
  Key: string;
  Value?: string;
}

/**
 * Optional shape for some SP profile responses
 */
export interface IUserProfileResponse {
  AccountName?: string;
  DisplayName?: string;
  Email?: string;
  Title?: string;
  UserProfileProperties?: IUserProfileProperty[];
}

/**
 * Your app-level user profile model
 */
export interface IUserProfile {
  /** Logged-in user basics (optional) */
  accountName?: string;
  displayName?: string;
  email?: string;

  /** Convenience property: resolved location (Office/Location) */
  Location?: string;

  /** Raw profile properties from SharePoint */
  UserProfileProperties: IUserProfileProperty[];

  /**
   * Helper function: get property by key (optional but useful)
   * Example: profile.get("SPS-Location")
   */
  get?: (key: string) => string;
}

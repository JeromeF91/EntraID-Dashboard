export interface SignInGeo {
  city?: string | null;
  state?: string | null;
  countryOrRegion?: string | null;
  geoCoordinates?: {
    latitude?: number | null;
    longitude?: number | null;
    altitude?: number | null;
  } | null;
}

export interface DeviceDetail {
  deviceId?: string | null;
  displayName?: string | null;
  operatingSystem?: string | null;
  browser?: string | null;
  isCompliant?: boolean | null;
  isManaged?: boolean | null;
  trustType?: string | null;
}

export interface NetworkLocationDetail {
  networkType?: string | null;
  networkNames?: string[] | null;
}

export interface SignInStatus {
  errorCode?: number | null;
  failureReason?: string | null;
  additionalDetails?: string | null;
}

export interface SignInRecord {
  id?: string;
  createdDateTime?: string;
  userDisplayName?: string | null;
  userPrincipalName?: string | null;
  userId?: string | null;
  ipAddress?: string | null;
  clientAppUsed?: string | null;
  appDisplayName?: string | null;
  isInteractive?: boolean | null;
  location?: SignInGeo | null;
  deviceDetail?: DeviceDetail | null;
  /** Conditional Access named locations resolved for this sign-in (trusted / other). */
  networkLocationDetails?: NetworkLocationDetail[] | null;
  /** Present on Graph sign-ins; `errorCode` 0 = success. */
  status?: SignInStatus | null;
}

export interface DaySummary {
  dateKey: string;
  firstUtc: string;
  lastUtc: string;
  signInCount: number;
  /** Sign-ins that matched a CA trusted named location (IP) for that day. */
  trustedNamedLocationCount: number;
  uniqueUsers: number;
  locationCounts: Record<string, number>;
}

export interface UserDayActivity {
  dateKey: string;
  firstUtc: string;
  lastUtc: string;
  signInCount: number;
  locations: Record<string, number>;
}

export interface UserSummary {
  userKey: string;
  userPrincipalName: string;
  userDisplayName: string;
  userId: string;
  totalSignIns: number;
  /** Distinct Eastern (America/New_York) calendar days with ≥1 sign-in where IP matched a CA trusted named location. */
  daysWithTrustedNamedLocation: number;
  countries: Record<string, number>;
  cities: Record<string, number>;
  firstSeenUtc: string;
  lastSeenUtc: string;
  byDay: UserDayActivity[];
}

export interface OutsideDesignatedRow {
  createdDateTime: string;
  userPrincipalName: string;
  userDisplayName: string;
  designatedCountry: string;
  actualCountry: string;
  city: string;
  ipAddress: string;
  appDisplayName: string;
  /** Heuristic from deviceDetail (joined/hybrid or managed+compliant). */
  trustedDevice: boolean;
  /** From networkLocationDetails when networkType is trustedNamedLocation. */
  trustedNamedLocation: boolean;
  /** Entra populated networkLocationDetails for this event. */
  hasNetworkLocationDetails: boolean;
  /** From `status.errorCode` (0 = success). */
  authKind: "success" | "failure" | "unknown";
  failureDetail: string;
}

export interface DashboardAggregate {
  recordCount: number;
  dateRange: { start: string; end: string } | null;
  byDay: DaySummary[];
  users: UserSummary[];
  parseError?: string;
}

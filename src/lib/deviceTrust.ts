import type { DeviceDetail, SignInRecord } from "./types";
import {
  assessNamedLocationTrust,
  type NamedLocationAssessment,
} from "./namedLocationTrust";
import {
  assessAuthenticationResult,
  type AuthenticationResultKind,
} from "./authStatus";

export interface DeviceTrustListingRow
  extends DeviceTrustAssessment, NamedLocationAssessment {
  createdDateTime: string;
  userPrincipalName: string;
  userDisplayName: string;
  appDisplayName: string;
  /** Client/public IP seen by Entra for this sign-in (`ipAddress`). */
  ipAddress: string;
  authKind: AuthenticationResultKind;
  failureDetail: string;
}

/** Matches Entra joined / hybrid corporate devices (case-insensitive). */
const JOINED_TRUST = new Set([
  "azure ad joined",
  "hybrid azure ad joined",
]);

export interface DeviceTrustAssessment {
  /** Has any device object with identity hints from Entra. */
  hasDeviceIdentity: boolean;
  /** Heuristic: joined/hybrid PC, or Intune-managed + compliant. */
  trusted: boolean;
  trustType: string;
  isManaged: boolean | null;
  isCompliant: boolean | null;
  deviceDisplayName: string;
}

export function assessDeviceTrust(rec: SignInRecord): DeviceTrustAssessment {
  const dd: DeviceDetail | null | undefined = rec.deviceDetail;
  const trustTypeRaw = (dd?.trustType ?? "").trim();
  const trustNorm = trustTypeRaw.toLowerCase();

  const nullDeviceId =
    !dd?.deviceId ||
    /^0{8}-0{4}-0{4}-0{4}-0{12}$/i.test(String(dd.deviceId).trim());

  const hasDeviceIdentity = !!(
    dd &&
    (trustTypeRaw !== "" ||
      (dd.displayName && dd.displayName.trim() !== "") ||
      !nullDeviceId)
  );

  const joined = JOINED_TRUST.has(trustNorm);
  const managed = dd?.isManaged === true ? true : dd?.isManaged === false ? false : null;
  const compliant =
    dd?.isCompliant === true ? true : dd?.isCompliant === false ? false : null;

  const trusted =
    joined ||
    (managed === true && compliant === true);

  return {
    hasDeviceIdentity,
    trusted,
    trustType: trustTypeRaw || "—",
    isManaged: managed,
    isCompliant: compliant,
    deviceDisplayName: (dd?.displayName ?? "").trim() || "—",
  };
}

export interface DeviceTrustTotals {
  /** Rows with a usable device identity in the log. */
  withDeviceIdentity: number;
  /** Subset that pass the trusted heuristic. */
  trustedDevice: number;
  /** Has device info but not classified as trusted. */
  notTrustedWithIdentity: number;
  /** No meaningful deviceDetail (unknown / browser-only / not sent). */
  noDeviceIdentity: number;
}

export function summarizeDeviceTrust(records: SignInRecord[]): DeviceTrustTotals {
  let withDeviceIdentity = 0;
  let trustedDevice = 0;
  let notTrustedWithIdentity = 0;
  let noDeviceIdentity = 0;

  for (const rec of records) {
    if (!rec.createdDateTime || Number.isNaN(new Date(rec.createdDateTime).getTime())) {
      continue;
    }
    const a = assessDeviceTrust(rec);
    if (!a.hasDeviceIdentity) {
      noDeviceIdentity += 1;
      continue;
    }
    withDeviceIdentity += 1;
    if (a.trusted) trustedDevice += 1;
    else notTrustedWithIdentity += 1;
  }

  return {
    withDeviceIdentity,
    trustedDevice,
    notTrustedWithIdentity,
    noDeviceIdentity,
  };
}

export function buildDeviceTrustListing(records: SignInRecord[]): DeviceTrustListingRow[] {
  const rows: DeviceTrustListingRow[] = [];
  for (const rec of records) {
    if (!rec.createdDateTime || Number.isNaN(new Date(rec.createdDateTime).getTime())) {
      continue;
    }
    const a = assessDeviceTrust(rec);
    const loc = assessNamedLocationTrust(rec);
    const auth = assessAuthenticationResult(rec);
    rows.push({
      ...a,
      ...loc,
      createdDateTime: rec.createdDateTime,
      userPrincipalName: rec.userPrincipalName?.trim() || "—",
      userDisplayName: rec.userDisplayName?.trim() || "—",
      appDisplayName: rec.appDisplayName?.trim() || "—",
      ipAddress: rec.ipAddress?.trim() || "—",
      authKind: auth.kind,
      failureDetail: auth.failureDetail,
    });
  }
  return rows.sort(
    (a, b) =>
      new Date(b.createdDateTime).getTime() - new Date(a.createdDateTime).getTime(),
  );
}

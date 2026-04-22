import type { SignInRecord } from "./types";

/** Graph enum value for IP matching a Conditional Access trusted named location. */
const TRUSTED_NAMED_LOCATION_TYPE = "trustednamedlocation";

export interface NamedLocationAssessment {
  /** Entra populated `networkLocationDetails` with at least one entry. */
  hasNetworkLocationDetails: boolean;
  /** Source IP mapped to a trusted named location (networkType trustedNamedLocation). */
  trustedNamedLocation: boolean;
  /** Named locations from CA (human-readable), comma-separated. */
  namedLocationLabels: string;
  /** Raw networkType strings from Entra (e.g. trustedNamedLocation). */
  networkTypesLabel: string;
}

function normalizeNetworkType(raw: unknown): string {
  return String(raw ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "");
}

export function assessNamedLocationTrust(rec: SignInRecord): NamedLocationAssessment {
  const raw = rec.networkLocationDetails;
  const list = Array.isArray(raw) ? raw : [];
  const hasNetworkLocationDetails = list.length > 0;

  const types = new Set<string>();
  const names = new Set<string>();

  for (const n of list) {
    const nt = (n?.networkType ?? "").trim();
    if (nt) types.add(nt);
    if (Array.isArray(n?.networkNames)) {
      for (const nn of n.networkNames) {
        if (typeof nn === "string" && nn.trim()) names.add(nn.trim());
      }
    }
  }

  const trustedNamedLocation = list.some(
    (n) => normalizeNetworkType(n?.networkType) === TRUSTED_NAMED_LOCATION_TYPE,
  );

  const namedLocationLabels = [...names].sort().join(", ") || "—";
  const networkTypesLabel =
    [...types].sort().join(", ") || (hasNetworkLocationDetails ? "—" : "");

  return {
    hasNetworkLocationDetails,
    trustedNamedLocation,
    namedLocationLabels,
    networkTypesLabel,
  };
}

export interface NamedLocationTotals {
  trustedNamedLocation: number;
  /** Has networkLocationDetails entries but none are trusted named. */
  networkDetailsOther: number;
  /** Empty or missing networkLocationDetails. */
  noNetworkDetails: number;
}

export function summarizeNamedLocationTrust(records: SignInRecord[]): NamedLocationTotals {
  let trustedNamedLocation = 0;
  let networkDetailsOther = 0;
  let noNetworkDetails = 0;

  for (const rec of records) {
    if (!rec.createdDateTime || Number.isNaN(new Date(rec.createdDateTime).getTime())) {
      continue;
    }
    const a = assessNamedLocationTrust(rec);
    if (!a.hasNetworkLocationDetails) {
      noNetworkDetails += 1;
    } else if (a.trustedNamedLocation) {
      trustedNamedLocation += 1;
    } else {
      networkDetailsOther += 1;
    }
  }

  return {
    trustedNamedLocation,
    networkDetailsOther,
    noNetworkDetails,
  };
}

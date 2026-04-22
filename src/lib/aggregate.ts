import type {
  DashboardAggregate,
  DaySummary,
  OutsideDesignatedRow,
  SignInRecord,
  UserDayActivity,
  UserSummary,
} from "./types";
import { assessAuthenticationResult } from "./authStatus";
import { easternDateKey } from "./easternTime";
import { assessDeviceTrust } from "./deviceTrust";
import { assessNamedLocationTrust } from "./namedLocationTrust";

function locationLabel(rec: SignInRecord): string {
  const loc = rec.location;
  if (!loc) return "Unknown";
  const parts = [
    loc.city,
    loc.state,
    loc.countryOrRegion,
  ].filter(Boolean) as string[];
  return parts.length ? parts.join(", ") : "Unknown";
}

function countryCode(rec: SignInRecord): string {
  const c = rec.location?.countryOrRegion?.trim();
  return c ? c.toUpperCase() : "";
}

function cityLabel(rec: SignInRecord): string {
  const city = rec.location?.city?.trim();
  return city || "";
}

function userKey(rec: SignInRecord): string {
  return (
    rec.userPrincipalName?.trim().toLowerCase() ||
    rec.userId?.trim().toLowerCase() ||
    "unknown"
  );
}

export function parseSignInJson(text: string): {
  records: SignInRecord[];
  error?: string;
} {
  try {
    const parsed = JSON.parse(text) as unknown;
    if (!Array.isArray(parsed)) {
      return { records: [], error: "JSON root must be an array of sign-in objects." };
    }
    return { records: parsed as SignInRecord[] };
  } catch (e) {
    const msg = e instanceof Error ? e.message : "Invalid JSON.";
    return { records: [], error: msg };
  }
}

export function aggregateSignIns(records: SignInRecord[]): DashboardAggregate {
  const withDates = records.filter(
    (r) => r.createdDateTime && !Number.isNaN(new Date(r.createdDateTime).getTime()),
  );
  const sorted = [...withDates].sort(
    (a, b) =>
      new Date(a.createdDateTime!).getTime() - new Date(b.createdDateTime!).getTime(),
  );

  if (sorted.length === 0) {
    return {
      recordCount: records.length,
      dateRange: null,
      byDay: [],
      users: [],
    };
  }

  const dayMap = new Map<
    string,
    {
      firstUtc: string;
      lastUtc: string;
      count: number;
      trustedNamedLocationCount: number;
      users: Set<string>;
      locationCounts: Record<string, number>;
    }
  >();

  const userMap = new Map<
    string,
    {
      userPrincipalName: string;
      userDisplayName: string;
      userId: string;
      totalSignIns: number;
      trustedNamedDays: Set<string>;
      countries: Record<string, number>;
      cities: Record<string, number>;
      firstSeenUtc: string;
      lastSeenUtc: string;
      days: Map<
        string,
        {
          firstUtc: string;
          lastUtc: string;
          count: number;
          locations: Record<string, number>;
        }
      >;
    }
  >();

  for (const rec of sorted) {
    const iso = rec.createdDateTime!;
    const dk = easternDateKey(iso);
    if (dk === "invalid") continue;

    const uk = userKey(rec);
    const locLabel = locationLabel(rec);
    const cc = countryCode(rec);
    const city = cityLabel(rec);
    const nlt = assessNamedLocationTrust(rec);

    let day = dayMap.get(dk);
    if (!day) {
      day = {
        firstUtc: iso,
        lastUtc: iso,
        count: 0,
        trustedNamedLocationCount: 0,
        users: new Set(),
        locationCounts: {},
      };
      dayMap.set(dk, day);
    }
    day.count += 1;
    if (nlt.trustedNamedLocation) {
      day.trustedNamedLocationCount += 1;
    }
    day.firstUtc = iso < day.firstUtc ? iso : day.firstUtc;
    day.lastUtc = iso > day.lastUtc ? iso : day.lastUtc;
    day.users.add(uk);
    day.locationCounts[locLabel] = (day.locationCounts[locLabel] ?? 0) + 1;

    let u = userMap.get(uk);
    if (!u) {
      u = {
        userPrincipalName: rec.userPrincipalName?.trim() || uk,
        userDisplayName: rec.userDisplayName?.trim() || uk,
        userId: rec.userId?.trim() || "",
        totalSignIns: 0,
        trustedNamedDays: new Set(),
        countries: {},
        cities: {},
        firstSeenUtc: iso,
        lastSeenUtc: iso,
        days: new Map(),
      };
      userMap.set(uk, u);
    }
    u.totalSignIns += 1;
    if (nlt.trustedNamedLocation) {
      u.trustedNamedDays.add(dk);
    }
    u.firstSeenUtc = iso < u.firstSeenUtc ? iso : u.firstSeenUtc;
    u.lastSeenUtc = iso > u.lastSeenUtc ? iso : u.lastSeenUtc;
    if (cc) {
      u.countries[cc] = (u.countries[cc] ?? 0) + 1;
    }
    if (city) {
      u.cities[city] = (u.cities[city] ?? 0) + 1;
    }

    let ud = u.days.get(dk);
    if (!ud) {
      ud = {
        firstUtc: iso,
        lastUtc: iso,
        count: 0,
        locations: {},
      };
      u.days.set(dk, ud);
    }
    ud.count += 1;
    ud.firstUtc = iso < ud.firstUtc ? iso : ud.firstUtc;
    ud.lastUtc = iso > ud.lastUtc ? iso : ud.lastUtc;
    ud.locations[locLabel] = (ud.locations[locLabel] ?? 0) + 1;
  }

  const byDay: DaySummary[] = [...dayMap.entries()]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([dateKey, v]) => ({
      dateKey,
      firstUtc: v.firstUtc,
      lastUtc: v.lastUtc,
      signInCount: v.count,
      trustedNamedLocationCount: v.trustedNamedLocationCount,
      uniqueUsers: v.users.size,
      locationCounts: v.locationCounts,
    }));

  const users: UserSummary[] = [...userMap.entries()]
    .map(([userKeyVal, u]) => {
      const byDayUser: UserDayActivity[] = [...u.days.entries()]
        .sort(([a], [b]) => a.localeCompare(b))
        .map(([dateKey, d]) => ({
          dateKey,
          firstUtc: d.firstUtc,
          lastUtc: d.lastUtc,
          signInCount: d.count,
          locations: d.locations,
        }));
      return {
        userKey: userKeyVal,
        userPrincipalName: u.userPrincipalName,
        userDisplayName: u.userDisplayName,
        userId: u.userId,
        totalSignIns: u.totalSignIns,
        daysWithTrustedNamedLocation: u.trustedNamedDays.size,
        countries: u.countries,
        cities: u.cities,
        firstSeenUtc: u.firstSeenUtc,
        lastSeenUtc: u.lastSeenUtc,
        byDay: byDayUser,
      };
    })
    .sort((a, b) => b.totalSignIns - a.totalSignIns);

  const start = sorted[0].createdDateTime!;
  const end = sorted[sorted.length - 1].createdDateTime!;

  return {
    recordCount: records.length,
    dateRange: { start, end },
    byDay,
    users,
  };
}

export function findOutsideDesignated(
  records: SignInRecord[],
  designatedByUpn: Record<string, string>,
): OutsideDesignatedRow[] {
  const designated = Object.fromEntries(
    Object.entries(designatedByUpn).map(([k, v]) => [
      k.trim().toLowerCase(),
      v.trim().toUpperCase(),
    ]),
  );

  const rows: OutsideDesignatedRow[] = [];

  for (const rec of records) {
    const upn = rec.userPrincipalName?.trim();
    if (!upn) continue;
    const expected = designated[upn.toLowerCase()];
    if (!expected) continue;

    const actual = countryCode(rec);
    if (!actual) continue;

    if (actual !== expected) {
      const locTrust = assessNamedLocationTrust(rec);
      const auth = assessAuthenticationResult(rec);
      rows.push({
        createdDateTime: rec.createdDateTime || "",
        userPrincipalName: upn,
        userDisplayName: rec.userDisplayName?.trim() || upn,
        designatedCountry: expected,
        actualCountry: actual,
        city: rec.location?.city?.trim() || "",
        ipAddress: rec.ipAddress?.trim() || "",
        appDisplayName: rec.appDisplayName?.trim() || "",
        trustedDevice: assessDeviceTrust(rec).trusted,
        trustedNamedLocation: locTrust.trustedNamedLocation,
        hasNetworkLocationDetails: locTrust.hasNetworkLocationDetails,
        authKind: auth.kind,
        failureDetail: auth.failureDetail,
      });
    }
  }

  return rows.sort(
    (a, b) =>
      new Date(b.createdDateTime).getTime() - new Date(a.createdDateTime).getTime(),
  );
}

export function inferMajorityCountry(user: UserSummary): string | null {
  const entries = Object.entries(user.countries);
  if (entries.length === 0) return null;
  entries.sort((a, b) => b[1] - a[1]);
  return entries[0][0];
}

import { useCallback, useMemo, useState } from "react";
import type { TdHTMLAttributes } from "react";
import {
  Bar,
  BarChart,
  CartesianGrid,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";
import {
  aggregateSignIns,
  findOutsideDesignated,
  inferMajorityCountry,
  parseSignInJson,
} from "./lib/aggregate";
import {
  buildDeviceTrustListing,
  summarizeDeviceTrust,
  type DeviceTrustListingRow,
} from "./lib/deviceTrust";
import { summarizeNamedLocationTrust } from "./lib/namedLocationTrust";
import type { AuthenticationResultKind } from "./lib/authStatus";
import {
  formatEasternDateOnly,
  formatEasternDateTime,
  formatEasternTimeOnly,
} from "./lib/easternTime";
import type {
  DashboardAggregate,
  OutsideDesignatedRow,
  SignInRecord,
  UserSummary,
} from "./lib/types";

const STORAGE_DESIGNATED = "entraid-dashboard-designated-json";

function loadDesignatedFromStorage(): string {
  try {
    return localStorage.getItem(STORAGE_DESIGNATED) ?? "";
  } catch {
    return "";
  }
}

/** Matches labels used in connection rows (`appDisplayName` or "—"). */
function recordAppLabel(rec: SignInRecord): string {
  return rec.appDisplayName?.trim() || "—";
}

function parseDesignatedJson(text: string): Record<string, string> {
  const t = text.trim();
  if (!t) return {};
  try {
    const o = JSON.parse(t) as unknown;
    if (o === null || typeof o !== "object" || Array.isArray(o)) return {};
    const out: Record<string, string> = {};
    for (const [k, v] of Object.entries(o)) {
      if (typeof v === "string" && k.trim()) {
        out[k.trim()] = v.trim().toUpperCase();
      }
    }
    return out;
  } catch {
    return {};
  }
}

export function App() {
  const [rawJson, setRawJson] = useState("");
  const [parsedRecords, setParsedRecords] = useState<SignInRecord[]>([]);
  const [recordsLength, setRecordsLength] = useState(0);
  const [applicationFilter, setApplicationFilter] = useState("");
  const [designatedText, setDesignatedText] = useState(loadDesignatedFromStorage);
  const [selectedUser, setSelectedUser] = useState<UserSummary | null>(null);
  const [userFilter, setUserFilter] = useState("");
  const [parseMessage, setParseMessage] = useState<string | null>(null);
  const [deviceTrustFilter, setDeviceTrustFilter] = useState<
    "all" | "trusted" | "notTrusted" | "noIdentity"
  >("all");
  const [namedLocationFilter, setNamedLocationFilter] = useState<
    "all" | "trustedNamed" | "detailsOther" | "noDetails"
  >("all");

  const designatedMap = useMemo(
    () => parseDesignatedJson(designatedText),
    [designatedText],
  );

  const applyJson = useCallback((text: string) => {
    setRawJson(text);
    const { records, error } = parseSignInJson(text);
    setParsedRecords(records);
    setRecordsLength(records.length);
    setApplicationFilter("");
    if (error) {
      setParseMessage(error);
      return;
    }
    setParseMessage(
      records.length === 0 ? "Array is empty." : null,
    );
  }, []);

  const applicationOptions = useMemo(() => {
    const counts = new Map<string, number>();
    for (const r of parsedRecords) {
      const label = recordAppLabel(r);
      counts.set(label, (counts.get(label) ?? 0) + 1);
    }
    return [...counts.entries()]
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([name, count]) => ({ name, count }));
  }, [parsedRecords]);

  const viewRecords = useMemo(() => {
    if (!applicationFilter) return parsedRecords;
    return parsedRecords.filter((r) => recordAppLabel(r) === applicationFilter);
  }, [parsedRecords, applicationFilter]);

  const aggregate: DashboardAggregate | null = useMemo(() => {
    if (viewRecords.length === 0) return null;
    return aggregateSignIns(viewRecords);
  }, [viewRecords]);

  const outsideRows: OutsideDesignatedRow[] = useMemo(() => {
    if (viewRecords.length === 0) return [];
    return findOutsideDesignated(viewRecords, designatedMap);
  }, [viewRecords, designatedMap]);

  const deviceTrustTotals = useMemo(
    () => summarizeDeviceTrust(viewRecords),
    [viewRecords],
  );

  const namedLocationTotals = useMemo(
    () => summarizeNamedLocationTrust(viewRecords),
    [viewRecords],
  );

  const deviceTrustRowsFull = useMemo(
    () => buildDeviceTrustListing(viewRecords),
    [viewRecords],
  );

  const deviceTrustRows = useMemo(() => {
    let rows = deviceTrustRowsFull;
    if (deviceTrustFilter !== "all") {
      rows = rows.filter((row) => {
        if (deviceTrustFilter === "trusted") return row.trusted && row.hasDeviceIdentity;
        if (deviceTrustFilter === "notTrusted") return row.hasDeviceIdentity && !row.trusted;
        return !row.hasDeviceIdentity;
      });
    }
    if (namedLocationFilter !== "all") {
      rows = rows.filter((row) => {
        if (namedLocationFilter === "trustedNamed") return row.trustedNamedLocation;
        if (namedLocationFilter === "detailsOther") {
          return row.hasNetworkLocationDetails && !row.trustedNamedLocation;
        }
        return !row.hasNetworkLocationDetails;
      });
    }
    return rows;
  }, [deviceTrustRowsFull, deviceTrustFilter, namedLocationFilter]);

  const chartData = useMemo(() => {
    if (!aggregate?.byDay.length) return [];
    return aggregate.byDay.map((d) => ({
      date: d.dateKey,
      signIns: d.signInCount,
      users: d.uniqueUsers,
    }));
  }, [aggregate]);

  const filteredUsers = useMemo(() => {
    if (!aggregate?.users) return [];
    const q = userFilter.trim().toLowerCase();
    if (!q) return aggregate.users;
    return aggregate.users.filter(
      (u) =>
        u.userPrincipalName.toLowerCase().includes(q) ||
        u.userDisplayName.toLowerCase().includes(q),
    );
  }, [aggregate, userFilter]);

  const persistDesignated = (next: string) => {
    setDesignatedText(next);
    try {
      localStorage.setItem(STORAGE_DESIGNATED, next);
    } catch {
      /* ignore */
    }
  };

  return (
    <div
      style={{
        maxWidth: 1280,
        margin: "0 auto",
        padding: "24px 20px 48px",
        width: "100%",
        boxSizing: "border-box",
        overflowX: "hidden",
      }}
    >
      <header style={{ marginBottom: 28 }}>
        <h1 style={{ margin: "0 0 8px", fontSize: "1.75rem", fontWeight: 600 }}>
          Entra ID interactive sign-ins
        </h1>
        <p style={{ margin: 0, color: "var(--muted)", maxWidth: 720 }}>
          Load exported sign-in JSON (array of records). The dashboard summarizes activity by day
          and location, flags sign-ins where geo country differs from your designated country per
          user (ISO 3166-1 alpha-2 — match Entra&apos;s{" "}
          <span className="mono">location.countryOrRegion</span>), and labels each connection as a{" "}
          <strong>trusted device</strong> when{" "}
          <span className="mono">deviceDetail.trustType</span> is Azure AD joined or hybrid Azure
          AD joined, or when the device is Intune-managed and compliant (
          <span className="mono">isManaged</span> &amp;{" "}
          <span className="mono">isCompliant</span>).{" "}
          <strong>Trusted named location</strong> uses Entra&apos;s{" "}
          <span className="mono">networkLocationDetails</span> when{" "}
          <span className="mono">networkType</span> is{" "}
          <span className="mono">trustedNamedLocation</span> (IP matched a Conditional Access
          trusted named location); <span className="mono">networkNames</span> lists the friendly
          names. Timestamps and per-day buckets use <strong>US Eastern</strong> (
          <span className="mono">America/New_York</span>, EST/EDT).
        </p>
      </header>

      <section
        style={{
          display: "grid",
          gap: 16,
          marginBottom: 28,
          padding: 20,
          background: "var(--surface)",
          border: "1px solid var(--border)",
          borderRadius: 12,
        }}
      >
        <div>
          <label style={{ display: "block", fontWeight: 600, marginBottom: 8 }}>
            Sign-in JSON
          </label>
          <textarea
            value={rawJson}
            onChange={(e) => applyJson(e.target.value)}
            placeholder='Paste exported JSON array, e.g. [ { "createdDateTime": "...", "userPrincipalName": "...", "location": { ... } }, ... ]'
            spellCheck={false}
            style={{
              width: "100%",
              minHeight: 140,
              padding: 12,
              borderRadius: 8,
              border: "1px solid var(--border)",
              background: "var(--bg)",
              color: "var(--text)",
              resize: "vertical",
            }}
          />
          <div style={{ marginTop: 8, display: "flex", flexWrap: "wrap", gap: 12, alignItems: "center" }}>
            <label style={{ cursor: "pointer", color: "var(--accent)" }}>
              <input
                type="file"
                accept=".json,application/json"
                style={{ display: "none" }}
                onChange={(e) => {
                  const f = e.target.files?.[0];
                  if (!f) return;
                  const reader = new FileReader();
                  reader.onload = () => {
                    if (typeof reader.result === "string") applyJson(reader.result);
                  };
                  reader.readAsText(f);
                  e.target.value = "";
                }}
              />
              Choose file
            </label>
            <span style={{ color: "var(--muted)", fontSize: 13 }}>
              {recordsLength > 0 ? `${recordsLength.toLocaleString()} rows parsed` : "No data"}
            </span>
            {parseMessage && (
              <span style={{ color: "var(--danger)", fontSize: 13 }}>{parseMessage}</span>
            )}
          </div>
        </div>

        <div>
          <label style={{ display: "block", fontWeight: 600, marginBottom: 8 }}>
            Designated country by UPN (JSON object)
          </label>
          <textarea
            value={designatedText}
            onChange={(e) => persistDesignated(e.target.value)}
            spellCheck={false}
            placeholder={'{\n  "jane.doe@contoso.com": "FR",\n  "john@contoso.com": "CA"\n}'}
            style={{
              width: "100%",
              minHeight: 88,
              padding: 12,
              borderRadius: 8,
              border: "1px solid var(--border)",
              background: "var(--bg)",
              color: "var(--text)",
              fontFamily: "var(--mono)",
              fontSize: 13,
              resize: "vertical",
            }}
          />
          <p style={{ margin: "8px 0 0", fontSize: 13, color: "var(--muted)" }}>
            Keys must match <span className="mono">userPrincipalName</span>. Values are expected
            countries (same codes as Entra geo). Users omitted here are not evaluated for
            out-of-location.
          </p>
        </div>
      </section>

      {parsedRecords.length > 0 && (
        <section
          style={{
            marginBottom: 28,
            padding: "14px 18px",
            background: "var(--surface)",
            border: "1px solid var(--border)",
            borderRadius: 12,
            display: "flex",
            flexWrap: "wrap",
            alignItems: "center",
            gap: 14,
          }}
        >
          <label htmlFor="app-filter" style={{ fontWeight: 600 }}>
            Application
          </label>
          <select
            id="app-filter"
            value={applicationFilter}
            onChange={(e) => setApplicationFilter(e.target.value)}
            style={{
              flex: "1 1 280px",
              minWidth: 220,
              maxWidth: "100%",
              padding: "8px 12px",
              borderRadius: 8,
              border: "1px solid var(--border)",
              background: "var(--bg)",
              color: "var(--text)",
            }}
          >
            <option value="">
              All applications ({parsedRecords.length.toLocaleString()} rows)
            </option>
            {applicationOptions.map(({ name, count }) => (
              <option key={name} value={name}>
                {name} ({count.toLocaleString()})
              </option>
            ))}
          </select>
          {applicationFilter !== "" && (
            <span style={{ fontSize: 13, color: "var(--muted)" }}>
              Showing {viewRecords.length.toLocaleString()} rows for this app (
              {parsedRecords.length.toLocaleString()} total loaded)
            </span>
          )}
        </section>
      )}

      {aggregate && aggregate.byDay.length > 0 && (
        <>
          <section style={{ marginBottom: 28 }}>
            <h2 style={{ fontSize: "1.15rem", margin: "0 0 12px" }}>Summary</h2>
            <div
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(auto-fill, minmax(200px, 1fr))",
                gap: 12,
              }}
            >
              <StatCard
                label="Records (with valid time)"
                value={aggregate.byDay.reduce((s, d) => s + d.signInCount, 0).toLocaleString()}
              />
              <StatCard
                label="Unique users"
                value={aggregate.users.length.toLocaleString()}
              />
              <StatCard
                label="Date span (Eastern)"
                value={
                  aggregate.dateRange
                    ? `${formatEasternDateOnly(aggregate.dateRange.start)} → ${formatEasternDateOnly(aggregate.dateRange.end)}`
                    : "—"
                }
              />
              <StatCard
                label="Outside designated"
                value={outsideRows.length.toLocaleString()}
                accent={outsideRows.length > 0 ? "danger" : "success"}
              />
              <StatCard
                label="Trusted device (heuristic)"
                value={deviceTrustTotals.trustedDevice.toLocaleString()}
              />
              <StatCard
                label="Known device, not trusted"
                value={deviceTrustTotals.notTrustedWithIdentity.toLocaleString()}
              />
              <StatCard
                label="No device identity in log"
                value={deviceTrustTotals.noDeviceIdentity.toLocaleString()}
              />
              <StatCard
                label="Trusted named location (IP)"
                value={namedLocationTotals.trustedNamedLocation.toLocaleString()}
              />
              <StatCard
                label="Named location info, not trusted"
                value={namedLocationTotals.networkDetailsOther.toLocaleString()}
              />
              <StatCard
                label="No networkLocationDetails in log"
                value={namedLocationTotals.noNetworkDetails.toLocaleString()}
              />
            </div>
          </section>

          <section style={{ marginBottom: 28 }}>
            <h2 style={{ fontSize: "1.15rem", margin: "0 0 12px" }}>
              Activity by day (Eastern dates)
            </h2>
            <div style={{ height: 280, width: "100%" }}>
              <ResponsiveContainer>
                <BarChart data={chartData} margin={{ top: 8, right: 8, left: 0, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#2d3a4d" />
                  <XAxis dataKey="date" tick={{ fill: "#8b9cb3", fontSize: 11 }} />
                  <YAxis tick={{ fill: "#8b9cb3", fontSize: 11 }} />
                  <Tooltip
                    contentStyle={{
                      background: "#1a2332",
                      border: "1px solid #2d3a4d",
                      borderRadius: 8,
                    }}
                    labelStyle={{ color: "#e7ecf3" }}
                  />
                  <Bar dataKey="signIns" name="Sign-ins" fill="#3d8bfd" radius={[4, 4, 0, 0]} />
                  <Bar dataKey="users" name="Distinct users" fill="#22c55e" radius={[4, 4, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </section>

          <section style={{ marginBottom: 28 }}>
            <h2 style={{ fontSize: "1.15rem", margin: "0 0 12px" }}>
              Device and trusted named location (every connection)
            </h2>
            <p style={{ margin: "0 0 12px", fontSize: 13, color: "var(--muted)" }}>
              Filters combine (AND): narrow by device trust and/or Conditional Access named
              location. Rows are newest first (cap 500 displayed).
            </p>
            <div style={{ marginBottom: 8 }}>
              <span style={{ fontSize: 12, color: "var(--muted)", marginRight: 8 }}>Device:</span>
              <span style={{ display: "inline-flex", flexWrap: "wrap", gap: 8 }}>
                {(
                  [
                    ["all", "All"],
                    ["trusted", "Trusted device"],
                    ["notTrusted", "Known device, not trusted"],
                    ["noIdentity", "No device identity"],
                  ] as const
                ).map(([key, label]) => (
                  <button
                    key={key}
                    type="button"
                    onClick={() =>
                      setDeviceTrustFilter(key as typeof deviceTrustFilter)
                    }
                    style={{
                      padding: "6px 12px",
                      borderRadius: 8,
                      border: "1px solid var(--border)",
                      background:
                        deviceTrustFilter === key ? "var(--accent)" : "var(--surface2)",
                      color: deviceTrustFilter === key ? "#fff" : "var(--text)",
                      cursor: "pointer",
                      fontSize: 13,
                    }}
                  >
                    {label}
                  </button>
                ))}
              </span>
            </div>
            <div style={{ marginBottom: 12 }}>
              <span style={{ fontSize: 12, color: "var(--muted)", marginRight: 8 }}>
                Named location:
              </span>
              <span style={{ display: "inline-flex", flexWrap: "wrap", gap: 8 }}>
                {(
                  [
                    ["all", "All"],
                    ["trustedNamed", "Trusted named location"],
                    ["detailsOther", "Other named location"],
                    ["noDetails", "No networkLocationDetails"],
                  ] as const
                ).map(([key, label]) => (
                  <button
                    key={key}
                    type="button"
                    onClick={() =>
                      setNamedLocationFilter(key as typeof namedLocationFilter)
                    }
                    style={{
                      padding: "6px 12px",
                      borderRadius: 8,
                      border: "1px solid var(--border)",
                      background:
                        namedLocationFilter === key ? "#22c55e" : "var(--surface2)",
                      color: namedLocationFilter === key ? "#0f1419" : "var(--text)",
                      cursor: "pointer",
                      fontSize: 13,
                    }}
                  >
                    {label}
                  </button>
                ))}
              </span>
            </div>
            <div className="table-wrap">
              <table className="table-responsive">
                <thead>
                  <tr>
                    <th>Time (Eastern)</th>
                    <th>Authentication</th>
                    <th>Public IP</th>
                    <th>Trusted device</th>
                    <th>Trusted named loc.</th>
                    <th>Named locations</th>
                    <th>Trust type</th>
                    <th>Managed</th>
                    <th>Compliant</th>
                    <th>Device</th>
                    <th>User</th>
                    <th>Application</th>
                  </tr>
                </thead>
                <tbody>
                  {deviceTrustRows.slice(0, 500).map((row, i) => (
                    <DeviceTrustTableRow key={`${row.createdDateTime}-${i}`} row={row} />
                  ))}
                </tbody>
              </table>
            </div>
            <p style={{ marginTop: 8, fontSize: 13, color: "var(--muted)" }}>
              Showing {Math.min(500, deviceTrustRows.length).toLocaleString()} of{" "}
              {deviceTrustRows.length.toLocaleString()} filtered rows (
              {deviceTrustRowsFull.length.toLocaleString()} total with valid time).
            </p>
          </section>

          <section style={{ marginBottom: 28 }}>
            <h2 style={{ fontSize: "1.15rem", margin: "0 0 12px" }}>
              Outside designated location
            </h2>
            {outsideRows.length === 0 ? (
              <p style={{ color: "var(--muted)", margin: 0 }}>
                No mismatches, or no designated mapping / no resolved country on events.
              </p>
            ) : (
              <div className="table-wrap">
                <table className="table-responsive">
                  <thead>
                    <tr>
                      <th>Time (Eastern)</th>
                      <th>Authentication</th>
                      <th>User</th>
                      <th>Trusted dev.</th>
                      <th>Trusted named loc.</th>
                      <th>Expected</th>
                      <th>Actual</th>
                      <th>City</th>
                      <th>Public IP</th>
                      <th>Application</th>
                    </tr>
                  </thead>
                  <tbody>
                    {outsideRows.slice(0, 500).map((r, i) => (
                      <tr
                        key={`${r.createdDateTime}-${r.userPrincipalName}-${i}`}
                        style={{
                          backgroundColor:
                            r.authKind === "failure"
                              ? "rgba(248, 113, 113, 0.12)"
                              : undefined,
                        }}
                      >
                        <RL label="Time (Eastern)" className="mono">
                          {formatEasternDateTime(r.createdDateTime)}
                        </RL>
                        <RL label="Authentication">
                          <AuthResultCell kind={r.authKind} detail={r.failureDetail} />
                        </RL>
                        <RL label="User">
                          <div>{r.userDisplayName}</div>
                          <div className="mono" style={{ color: "var(--muted)" }}>
                            {r.userPrincipalName}
                          </div>
                        </RL>
                        <RL label="Trusted dev.">
                          <TrustedBadge trusted={r.trustedDevice} />
                        </RL>
                        <RL label="Trusted named loc.">
                          <NamedLocationBadge
                            hasDetails={r.hasNetworkLocationDetails}
                            trustedNamed={r.trustedNamedLocation}
                          />
                        </RL>
                        <RL label="Expected" className="mono">
                          {r.designatedCountry}
                        </RL>
                        <RL label="Actual" className="mono">
                          {r.actualCountry}
                        </RL>
                        <RL label="City">{r.city || "—"}</RL>
                        <RL label="Public IP" className="mono">
                          {r.ipAddress || "—"}
                        </RL>
                        <RL label="Application">{r.appDisplayName || "—"}</RL>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
            {outsideRows.length > 500 && (
              <p style={{ fontSize: 13, color: "var(--muted)", marginTop: 8 }}>
                Showing first 500 of {outsideRows.length} rows.
              </p>
            )}
          </section>

          <section style={{ marginBottom: 28 }}>
            <h2 style={{ fontSize: "1.15rem", margin: "0 0 12px" }}>Per day (Eastern)</h2>
            <div className="table-wrap">
              <table className="table-responsive">
                <thead>
                  <tr>
                    <th>Date</th>
                    <th>First sign-in</th>
                    <th>Last sign-in</th>
                    <th>Sign-ins</th>
                    <th>From trusted named loc.</th>
                    <th>Users</th>
                    <th>Top locations</th>
                  </tr>
                </thead>
                <tbody>
                  {aggregate.byDay.map((d) => (
                    <tr key={d.dateKey}>
                      <RL label="Date" className="mono">
                        {d.dateKey}
                      </RL>
                      <RL label="First sign-in (Eastern)" className="mono">
                        {formatEasternTimeOnly(d.firstUtc)}
                      </RL>
                      <RL label="Last sign-in (Eastern)" className="mono">
                        {formatEasternTimeOnly(d.lastUtc)}
                      </RL>
                      <RL label="Sign-ins">{d.signInCount.toLocaleString()}</RL>
                      <RL label="From trusted named loc.">
                        {d.trustedNamedLocationCount.toLocaleString()}
                      </RL>
                      <RL label="Users">{d.uniqueUsers}</RL>
                      <RL label="Top locations" style={{ maxWidth: "none" }}>
                        {topEntries(d.locationCounts, 4).join(" · ") || "—"}
                      </RL>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>

          <section>
            <h2 style={{ fontSize: "1.15rem", margin: "0 0 12px" }}>Users</h2>
            <input
              type="search"
              placeholder="Filter by name or UPN…"
              value={userFilter}
              onChange={(e) => setUserFilter(e.target.value)}
              style={{
                marginBottom: 12,
                padding: "8px 12px",
                width: "100%",
                maxWidth: 360,
                borderRadius: 8,
                border: "1px solid var(--border)",
                background: "var(--bg)",
                color: "var(--text)",
              }}
            />
            <div className="table-wrap" style={{ maxHeight: 360 }}>
              <table className="table-responsive">
                <thead>
                  <tr>
                    <th>User</th>
                    <th>Sign-ins</th>
                    <th>Days (trusted named loc.)</th>
                    <th>Majority country</th>
                    <th>Countries</th>
                    <th>Activity</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredUsers.map((u) => (
                    <tr
                      key={u.userKey}
                      style={{ cursor: "pointer" }}
                      onClick={() => setSelectedUser(u)}
                    >
                      <RL label="User">
                        <div>{u.userDisplayName}</div>
                        <div className="mono" style={{ color: "var(--muted)" }}>
                          {u.userPrincipalName}
                        </div>
                      </RL>
                      <RL label="Sign-ins">{u.totalSignIns.toLocaleString()}</RL>
                      <RL label="Days (trusted named loc.)">
                        {u.daysWithTrustedNamedLocation.toLocaleString()}
                      </RL>
                      <RL label="Majority country" className="mono">
                        {inferMajorityCountry(u) ?? "—"}
                      </RL>
                      <RL label="Countries" style={{ fontSize: 12 }}>
                        {formatCounts(u.countries)}
                      </RL>
                      <RL label="Activity">
                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            setSelectedUser(u);
                          }}
                          style={{
                            background: "var(--surface2)",
                            border: "1px solid var(--border)",
                            color: "var(--accent)",
                            borderRadius: 6,
                            padding: "4px 10px",
                            cursor: "pointer",
                          }}
                        >
                          Day breakdown
                        </button>
                      </RL>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>
        </>
      )}

      {selectedUser && (
        <UserModal user={selectedUser} onClose={() => setSelectedUser(null)} />
      )}
    </div>
  );
}

/** Table cell with `data-label` for stacked (“card”) layout on narrow viewports. */
function RL({
  label,
  children,
  ...props
}: { label: string } & TdHTMLAttributes<HTMLTableCellElement>) {
  return (
    <td data-label={label} {...props}>
      {children}
    </td>
  );
}

function TrustedBadge({ trusted }: { trusted: boolean }) {
  return (
    <span
      style={{
        display: "inline-block",
        padding: "2px 8px",
        borderRadius: 6,
        fontSize: 12,
        fontWeight: 600,
        background: trusted ? "rgba(74, 222, 128, 0.15)" : "rgba(248, 113, 113, 0.12)",
        color: trusted ? "var(--success)" : "var(--danger)",
      }}
    >
      {trusted ? "Yes" : "No"}
    </span>
  );
}

function NamedLocationBadge({
  hasDetails,
  trustedNamed,
}: {
  hasDetails: boolean;
  trustedNamed: boolean;
}) {
  if (!hasDetails) {
    return <span style={{ color: "var(--muted)" }}>—</span>;
  }
  return <TrustedBadge trusted={trustedNamed} />;
}

function AuthResultCell({
  kind,
  detail,
}: {
  kind: AuthenticationResultKind;
  detail: string;
}) {
  if (kind === "failure") {
    return (
      <span style={{ color: "var(--danger)", fontWeight: 700 }} title={detail || undefined}>
        Failed
      </span>
    );
  }
  if (kind === "unknown") {
    return (
      <span style={{ color: "var(--muted)" }} title="No status.errorCode on this row">
        Unknown
      </span>
    );
  }
  return <span style={{ color: "var(--muted)" }}>Success</span>;
}

function DeviceTrustTableRow({ row }: { row: DeviceTrustListingRow }) {
  const trustedShow =
    row.hasDeviceIdentity && row.trusted
      ? true
      : row.hasDeviceIdentity
        ? false
        : null;
  return (
    <tr
      style={{
        backgroundColor:
          row.authKind === "failure" ? "rgba(248, 113, 113, 0.12)" : undefined,
      }}
    >
      <RL label="Time (Eastern)" className="mono">
        {formatEasternDateTime(row.createdDateTime)}
      </RL>
      <RL label="Authentication">
        <AuthResultCell kind={row.authKind} detail={row.failureDetail} />
      </RL>
      <RL label="Public IP" className="mono">
        {row.ipAddress}
      </RL>
      <RL label="Trusted device">
        {trustedShow === null ? (
          <span style={{ color: "var(--muted)" }}>—</span>
        ) : (
          <TrustedBadge trusted={trustedShow} />
        )}
      </RL>
      <RL label="Trusted named loc.">
        <NamedLocationBadge
          hasDetails={row.hasNetworkLocationDetails}
          trustedNamed={row.trustedNamedLocation}
        />
      </RL>
      <RL
        label="Named locations"
        style={{ fontSize: 12, maxWidth: "none" }}
        title={row.namedLocationLabels !== "—" ? row.namedLocationLabels : undefined}
      >
        {row.namedLocationLabels}
      </RL>
      <RL label="Trust type" style={{ fontSize: 12 }}>
        {row.trustType}
      </RL>
      <RL label="Managed">{tri(row.isManaged)}</RL>
      <RL label="Compliant">{tri(row.isCompliant)}</RL>
      <RL label="Device" style={{ fontSize: 12 }}>
        {row.deviceDisplayName}
      </RL>
      <RL label="User">
        <div>{row.userDisplayName}</div>
        <div className="mono" style={{ color: "var(--muted)", fontSize: 12 }}>
          {row.userPrincipalName}
        </div>
      </RL>
      <RL label="Application" style={{ fontSize: 12 }}>
        {row.appDisplayName}
      </RL>
    </tr>
  );
}

function tri(v: boolean | null): string {
  if (v === true) return "Yes";
  if (v === false) return "No";
  return "—";
}

function StatCard({
  label,
  value,
  accent,
}: {
  label: string;
  value: string;
  accent?: "danger" | "success";
}) {
  const color =
    accent === "danger"
      ? "var(--danger)"
      : accent === "success"
        ? "var(--success)"
        : "var(--text)";
  return (
    <div
      style={{
        padding: "14px 16px",
        background: "var(--bg)",
        border: "1px solid var(--border)",
        borderRadius: 10,
      }}
    >
      <div style={{ fontSize: 12, color: "var(--muted)", marginBottom: 4 }}>{label}</div>
      <div style={{ fontSize: "1.35rem", fontWeight: 600, color }}>{value}</div>
    </div>
  );
}

function topEntries(counts: Record<string, number>, n: number): string[] {
  return Object.entries(counts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, n)
    .map(([k, v]) => `${k} (${v})`);
}

function formatCounts(counts: Record<string, number>): string {
  return Object.entries(counts)
    .sort((a, b) => b[1] - a[1])
    .map(([k, v]) => `${k}:${v}`)
    .join(", ");
}

function UserModal({ user, onClose }: { user: UserSummary; onClose: () => void }) {
  return (
    <div
      role="presentation"
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(0,0,0,0.65)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        padding: 16,
        zIndex: 50,
      }}
      onClick={onClose}
    >
      <div
        role="dialog"
        aria-modal="true"
        style={{
          width: "min(900px, 100%)",
          maxHeight: "90vh",
          overflow: "auto",
          background: "var(--surface)",
          border: "1px solid var(--border)",
          borderRadius: 12,
          padding: 24,
        }}
        onClick={(e) => e.stopPropagation()}
      >
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "start", gap: 16 }}>
          <div>
            <h3 style={{ margin: "0 0 4px", fontSize: "1.2rem" }}>{user.userDisplayName}</h3>
            <div className="mono" style={{ color: "var(--muted)", fontSize: 13 }}>
              {user.userPrincipalName}
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            style={{
              background: "var(--surface2)",
              border: "1px solid var(--border)",
              color: "var(--text)",
              borderRadius: 8,
              padding: "6px 14px",
              cursor: "pointer",
            }}
          >
            Close
          </button>
        </div>
        <p style={{ fontSize: 13, color: "var(--muted)", marginTop: 12 }}>
          First seen {formatEasternDateTime(user.firstSeenUtc)} · Last seen{" "}
          {formatEasternDateTime(user.lastSeenUtc)}
          <br />
          {user.daysWithTrustedNamedLocation.toLocaleString()} distinct Eastern calendar{" "}
          {user.daysWithTrustedNamedLocation === 1 ? "day" : "days"} with at least one sign-in from
          a trusted named location (Conditional Access).
        </p>
        <h4 style={{ margin: "20px 0 10px", fontSize: "1rem" }}>Per day (Eastern)</h4>
        <div className="table-wrap" style={{ maxHeight: 320 }}>
          <table className="table-responsive">
            <thead>
              <tr>
                <th>Date</th>
                <th>First</th>
                <th>Last</th>
                <th>Sign-ins</th>
                <th>Locations</th>
              </tr>
            </thead>
            <tbody>
              {user.byDay.map((d) => (
                <tr key={d.dateKey}>
                  <RL label="Date" className="mono">
                    {d.dateKey}
                  </RL>
                  <RL label="First (Eastern)" className="mono">
                    {formatEasternTimeOnly(d.firstUtc)}
                  </RL>
                  <RL label="Last (Eastern)" className="mono">
                    {formatEasternTimeOnly(d.lastUtc)}
                  </RL>
                  <RL label="Sign-ins">{d.signInCount}</RL>
                  <RL label="Locations" style={{ fontSize: 12 }}>
                    {topEntries(d.locations, 6).join(" · ")}
                  </RL>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

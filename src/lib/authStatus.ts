import type { SignInRecord } from "./types";

/** Derived from `status.errorCode`: 0 = success per Entra sign-in logs. */
export type AuthenticationResultKind = "success" | "failure" | "unknown";

export function assessAuthenticationResult(rec: SignInRecord): {
  kind: AuthenticationResultKind;
  failureDetail: string;
} {
  const raw = rec.status?.errorCode;
  if (typeof raw !== "number") {
    return { kind: "unknown", failureDetail: "" };
  }
  if (raw !== 0) {
    const parts = [rec.status?.failureReason, rec.status?.additionalDetails].filter(
      (x): x is string => typeof x === "string" && x.trim().length > 0,
    );
    return { kind: "failure", failureDetail: parts.join(" — ") };
  }
  return { kind: "success", failureDetail: "" };
}

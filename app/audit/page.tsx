import type { Metadata } from "next";
import AuditEditor from "./audit-editor";
export const metadata: Metadata = {
  title: "Audit workbook | FileBridge",
  alternates: { canonical: "/audit" },
};
export default function AuditPage() {
  return <AuditEditor />;
}

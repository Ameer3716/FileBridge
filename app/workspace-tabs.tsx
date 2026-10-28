"use client";
import Link from "next/link";
import { usePathname } from "next/navigation";
export default function WorkspaceTabs() {
  const path = usePathname();
  return (
    <nav className="workspace-tabs" aria-label="FileBridge tools">
      <Link href="/" aria-current={path === "/" ? "page" : undefined}>
        File transfer
      </Link>
      <Link href="/audit" aria-current={path === "/audit" ? "page" : undefined}>
        Audit workbook
      </Link>
    </nav>
  );
}

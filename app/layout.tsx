import type { Metadata } from "next";
import "./globals.css";
import "./audit.css";

export const metadata: Metadata = {
  title: "TMBill Revenue & KOT Audit | Morbido Express",
  description: "Developer-facing sales, VAT, payment, item, KOT, cancellation and complimentary audit dashboard.",
  openGraph: {
    title: "TMBill Revenue & KOT Audit",
    description: "Sales · VAT · Payments · KOT Controls",
    images: ["/tmbill-audit-social.png"],
  },
};

export default function RootLayout({children}:{children:React.ReactNode}) {
  return <html lang="en"><body>{children}</body></html>;
}

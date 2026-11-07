
import type { Metadata } from "next";
import "./globals.css";
import { ThemeProvider } from "./components/ThemeProvider";

export const metadata: Metadata = {
  title: "eSparks Document Conversion",
  description: "Simple and secure document conversion tools."
};

export default function RootLayout({
  children
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en" suppressHydrationWarning>
      <body>
        <ThemeProvider>
          {children}
        </ThemeProvider>
      </body>
    </html>
  );
}

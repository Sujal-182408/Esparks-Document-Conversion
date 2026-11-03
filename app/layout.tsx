import "./globals.css";

export const metadata = {
  title: "eSparks Document Conversion | Convert Documents Online",
  description:
    "Convert Word, PDF, PowerPoint and Markdown documents quickly and securely with eSparks Document Conversion.",
  icons: {
    icon: "/favicon.ico",
  },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}

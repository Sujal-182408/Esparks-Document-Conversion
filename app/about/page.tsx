import Navbar from "../components/Navbar";
import About from "../components/About";
import Footer from "../components/Footer";

export default function AboutPage() {
  return (
    <main className="min-h-screen bg-[#070b14] text-white">
      <Navbar />
      <About />
      <Footer />
    </main>
  );
}

"use client";

import {
  createContext,
  useContext,
  useEffect,
  useState,
  type ReactNode
} from "react";

type Theme = "dark" | "light";

type ThemeContextType = {
  theme: Theme;
  toggleTheme: () => void;
};

const ThemeContext = createContext<ThemeContextType | undefined>(undefined);

export function ThemeProvider({ children }: { children: ReactNode }) {
  const [theme, setTheme] = useState<Theme>("dark");

  useEffect(() => {
    const savedTheme = localStorage.getItem("esparks-theme");
    const initialTheme: Theme = savedTheme === "light" ? "light" : "dark";

    setTheme(initialTheme);
    document.documentElement.setAttribute("data-theme", initialTheme);
  }, []);

  const toggleTheme = () => {
    setTheme((currentTheme) => {
      const nextTheme: Theme = currentTheme === "dark" ? "light" : "dark";

      document.documentElement.setAttribute("data-theme", nextTheme);
      localStorage.setItem("esparks-theme", nextTheme);

      return nextTheme;
    });
  };

  return (
    <ThemeContext.Provider value={{ theme, toggleTheme }}>
      {children}
    </ThemeContext.Provider>
  );
}

export function useTheme() {
  const context = useContext(ThemeContext);

  if (!context) {
    throw new Error("useTheme must be used within ThemeProvider");
  }

  return context;
}

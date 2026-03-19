import "./components/xlsx-parser-demo";

const app = document.getElementById("app");

if (app) {
  app.innerHTML = `
    <main style="padding: 24px; max-width: 1440px; margin: 0 auto;">
      <header style="margin-bottom: 20px; font-family: 'Segoe UI', sans-serif;">
        <h1 style="margin: 0 0 8px; font-size: 28px;">XlsxParser Demo</h1>
        <p style="margin: 0; color: #5d6878;">
          Local Vite playground for workbook parsing and HTML rendering.
        </p>
      </header>

      <xlsx-parser-demo></xlsx-parser-demo>
    </main>
  `;
}

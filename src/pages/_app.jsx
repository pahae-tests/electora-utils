// import "@/styles/globals.css";
import Sidebar from "../components/Sidebar";

export default function App({
  Component,
  pageProps,
}) {
  return (
    <div className="app">

      <Sidebar />

      <main className="mainContent">
        <Component {...pageProps} />
      </main>

      <style jsx global>{`

        * {
          box-sizing: border-box;
        }

        html,
        body {
          margin: 0;
          padding: 0;

          min-height: 100%;

          background: #f4f2ee;
        }

        body {
          font-family:
            -apple-system,
            BlinkMacSystemFont,
            "Segoe UI",
            Roboto,
            Helvetica,
            Arial,
            sans-serif;

          color: #1e2124;
        }

        button,
        input,
        select,
        textarea {
          font-family: inherit;
        }

        button {
          -webkit-tap-highlight-color: transparent;
        }

        a {
          -webkit-tap-highlight-color: transparent;
        }

        /* =====================================================
           APP LAYOUT
        ===================================================== */

        .app {
          min-height: 100vh;
        }

        .mainContent {
          min-height: 100vh;

          margin-left: 270px;

          width: calc(
            100% - 270px
          );
        }

        /* =====================================================
           TABLET
        ===================================================== */

        @media (max-width: 900px) {

          .mainContent {
            margin-left: 250px;

            width: calc(
              100% - 250px
            );
          }

        }

        /* =====================================================
           MOBILE
        ===================================================== */

        @media (max-width: 700px) {

          .mainContent {
            margin-left: 0;

            width: 100%;

            padding-top: 0;
          }

        }

      `}</style>

    </div>
  );
}
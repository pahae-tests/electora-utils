import { useEffect, useState } from "react";
import Link from "next/link";
import { useRouter } from "next/router";
import {
    Search,
    RefreshCw,
    BarChart3,
    Users,
    ShieldCheck,
    ChevronRight,
    Menu,
    X,
    Database,
    Activity,
    FileSpreadsheet,
    GitMerge,
} from "lucide-react";

const menuItems = [
    {
        name: "Fetch",
        description: "Recherche intelligente dans les données de Electora",
        href: "/fetch",
        icon: RefreshCw,
    },
    {
        name: "Count",
        description: "Statistiques pour un parrain",
        href: "/count",
        icon: BarChart3,
    },
    {
        name: "Social",
        description: "Génère excel Moltaqa et excel Social pour chaque parrain",
        href: "/social",
        icon: Users,
    },
    {
        name: "Verify",
        description: "Génère un excel pour les électeurs Moltaqa non verifiés",
        href: "/verify",
        icon: ShieldCheck,
    },
    {
        name: "Merge",
        description: "Fusionne deux excels",
        href: "/merge",
        icon: GitMerge,
    },
    {
        name: "Excels",
        description: "Télécharge un excel pour chaque parrain",
        href: "/excels",
        icon: FileSpreadsheet,
    },
    {
        name: "Check",
        description: "Recherche rapide dans les électeurs d'un parrain",
        href: "/check",
        icon: Search,
    },
];

export default function Sidebar() {
    const router = useRouter();
    const [open, setOpen] = useState(false);

    useEffect(() => {
        setOpen(false);
    }, [router.pathname]);

    useEffect(() => {
        document.body.style.overflow = open ? "hidden" : "";

        return () => {
            document.body.style.overflow = "";
        };
    }, [open]);

    return (
        <>
            {/* =====================================================
          MOBILE BUTTON
      ===================================================== */}

            <button
                type="button"
                className="electoraMobileButton"
                onClick={() => setOpen(true)}
                aria-label="Ouvrir le menu"
            >
                <Menu size={21} strokeWidth={1.8} />
            </button>

            {/* =====================================================
          OVERLAY
      ===================================================== */}

            {open && (
                <div
                    className="electoraOverlay"
                    onClick={() => setOpen(false)}
                />
            )}

            {/* =====================================================
          SIDEBAR
      ===================================================== */}

            <aside
                className={`electoraSidebar ${open ? "electoraSidebarOpen" : ""
                    }`}
            >
                {/* ===================================================
            BRAND
        =================================================== */}

                <div className="electoraBrand">

                    <div className="electoraBrandLogo">
                        <Database
                            size={22}
                            strokeWidth={2}
                        />
                    </div>

                    <div className="electoraBrandInfo">

                        <div className="electoraBrandName">
                            Pahae Utils
                        </div>

                        <div className="electoraBrandTagline">
                            Pour Electora
                        </div>

                    </div>

                    <button
                        type="button"
                        className="electoraCloseButton"
                        onClick={() => setOpen(false)}
                        aria-label="Fermer"
                    >
                        <X size={21} />
                    </button>

                </div>

                {/* ===================================================
            SYSTEM STATUS
        =================================================== */}

                {/* <div className="electoraSystemStatus">

                    <div className="electoraStatusIcon">
                        <Activity
                            size={15}
                            strokeWidth={2}
                        />
                    </div>

                    <div className="electoraStatusInfo">

                        <span className="electoraStatusLabel">
                            SYSTÈME
                        </span>

                        <span className="electoraStatusValue">
                            Opérationnel
                        </span>

                    </div>

                    <span className="electoraStatusDot" />

                </div> */}

                {/* ===================================================
            NAVIGATION
        =================================================== */}

                <nav className="electoraNavigation">

                    <div className="electoraNavSectionTitle">
                        OUTILS
                    </div>

                    <div className="electoraNavList">

                        {menuItems.map((item) => {

                            const Icon = item.icon;

                            const isActive =
                                router.pathname === item.href ||
                                router.pathname.startsWith(
                                    item.href + "/"
                                );

                            return (
                                <Link
                                    key={item.href}
                                    href={item.href}
                                    className={`electoraNavItem ${isActive
                                            ? "electoraNavItemActive"
                                            : ""
                                        }`}
                                >

                                    <div className="electoraNavIcon">
                                        <Icon
                                            size={18}
                                            strokeWidth={
                                                isActive ? 2.1 : 1.8
                                            }
                                        />
                                    </div>

                                    <div className="electoraNavText">

                                        <span className="electoraNavName">
                                            {item.name}
                                        </span>

                                        <span className="electoraNavDescription">
                                            {item.description}
                                        </span>

                                    </div>

                                    <ChevronRight
                                        className="electoraNavArrow"
                                        size={15}
                                        strokeWidth={1.8}
                                    />

                                </Link>
                            );
                        })}

                    </div>

                </nav>

            </aside>

            <style jsx global>{`

        /* =====================================================
           SIDEBAR
        ===================================================== */

        .electoraSidebar {
          position: fixed !important;

          top: 0 !important;
          left: 0 !important;
          bottom: 0 !important;

          width: 278px !important;

          display: flex !important;
          flex-direction: column !important;

          background: #ffffff !important;

          border-right: 1px solid #e5e2dc !important;

          z-index: 1000 !important;

          box-shadow:
            2px 0 15px
            rgba(20, 25, 30, 0.025);

          overflow: hidden;
        }

        /* =====================================================
           BRAND
        ===================================================== */

        .electoraBrand {
          height: 82px;

          min-height: 82px;

          display: flex;

          align-items: center;

          padding: 0 20px;

          border-bottom: 1px solid #ece9e3;

          flex-shrink: 0;
        }

        .electoraBrandLogo {
          width: 42px;
          height: 42px;

          flex-shrink: 0;

          display: flex;

          align-items: center;
          justify-content: center;

          background: #17395f;

          color: #ffffff;

          border-radius: 7px;

          box-shadow:
            0 4px 10px
            rgba(23, 57, 95, 0.16);
        }

        .electoraBrandInfo {
          margin-left: 13px;

          min-width: 0;
        }

        .electoraBrandName {
          font-family:
            Georgia,
            "Times New Roman",
            serif;

          font-size: 25px;

          line-height: 1;

          font-weight: 600;

          color: #17202a;
        }

        .electoraBrandTagline {
          margin-top: 7px;

          color: #9a948b;

          font-size: 10px;

          font-weight: 700;

          letter-spacing: 0.07em;
        }

        .electoraCloseButton {
          display: none;

          margin-left: auto;

          width: 34px;
          height: 34px;

          align-items: center;
          justify-content: center;

          border: 0 !important;

          background: transparent !important;

          color: #6d6860;

          cursor: pointer;

          border-radius: 5px;
        }

        /* =====================================================
           STATUS
        ===================================================== */

        .electoraSystemStatus {
          margin: 27px 16px 0;

          padding: 12px 12px;

          display: flex;

          align-items: center;

          background: #f7f8f8;

          border: 1px solid #e5e8e6;

          border-radius: 8px;

          flex-shrink: 0;
        }

        .electoraStatusIcon {
          width: 32px;
          height: 32px;

          display: flex;

          align-items: center;
          justify-content: center;

          background: #ffffff;

          border: 1px solid #dfe4e1;

          border-radius: 6px;

          color: #4d8662;
        }

        .electoraStatusInfo {
          margin-left: 10px;

          display: flex;

          flex-direction: column;
        }

        .electoraStatusLabel {
          font-size: 9px;

          font-weight: 700;

          letter-spacing: 0.08em;

          color: #aaa49b;
        }

        .electoraStatusValue {
          margin-top: 3px;

          font-size: 13px;

          font-weight: 600;

          color: #555b58;
        }

        .electoraStatusDot {
          width: 8px;
          height: 8px;

          margin-left: auto;

          background: #4d8662;

          border-radius: 50%;

          box-shadow:
            0 0 0 4px
            rgba(77, 134, 98, 0.09);
        }

        /* =====================================================
           NAVIGATION
        ===================================================== */

        .electoraNavigation {
          flex: 1;

          min-height: 0;

          padding: 26px 12px 15px;

          overflow-y: auto;

          overflow-x: hidden;
        }

        .electoraNavSectionTitle {
          padding: 0 12px;

          margin-bottom: 12px;

          color: #999289;

          font-size: 10px;

          font-weight: 700;

          letter-spacing: 0.09em;
        }

        .electoraNavList {
          display: flex;

          flex-direction: column;

          gap: 4px;

          width: 100%;
        }

        /* =====================================================
           IMPORTANT :
           STYLE DU LINK
        ===================================================== */

        .electoraSidebar a.electoraNavItem {
          position: relative !important;

          width: 100% !important;

          min-height: 62px !important;

          display: flex !important;

          flex-direction: row !important;

          align-items: center !important;

          justify-content: flex-start !important;

          padding: 8px 10px !important;

          margin: 0 !important;

          box-sizing: border-box !important;

          text-decoration: none !important;

          color: #68625a !important;

          background: transparent !important;

          border: 1px solid transparent !important;

          border-radius: 7px !important;

          cursor: pointer !important;

          transition:
            background 0.16s ease,
            border-color 0.16s ease,
            color 0.16s ease;
        }

        .electoraSidebar a.electoraNavItem:hover {
          background: #f7f6f3 !important;

          border-color: #ece9e3 !important;

          color: #263e58 !important;
        }

        .electoraSidebar a.electoraNavItem.electoraNavItemActive {
          background: #f1f5f8 !important;

          border-color: #dce5ec !important;

          color: #17395f !important;
        }

        .electoraNavItemActive::before {
          content: "";

          position: absolute;

          left: -1px;

          top: 11px;
          bottom: 11px;

          width: 3px;

          background: #b18a4a;

          border-radius:
            0 3px 3px 0;
        }

        /* =====================================================
           ICON
        ===================================================== */

        .electoraNavIcon {
          width: 39px;
          height: 39px;

          min-width: 39px;

          flex-shrink: 0;

          display: flex;

          align-items: center;
          justify-content: center;

          background: #ffffff;

          border: 1px solid #e4e1db;

          border-radius: 6px;

          color: #79736a;
        }

        .electoraNavItem:hover
        .electoraNavIcon {
          color: #263e58;

          border-color: #d8dfe4;
        }

        .electoraNavItemActive
        .electoraNavIcon {
          color: #17395f;

          border-color: #cbd9e4;

          background: #ffffff;
        }

        /* =====================================================
           TEXT
        ===================================================== */

        .electoraNavText {
          min-width: 0;

          margin-left: 12px;

          display: flex;

          flex-direction: column;

          justify-content: center;

          flex: 1;
        }

        .electoraNavName {
          display: block;

          font-size: 13px;

          line-height: 1.2;

          font-weight: 650;

          color: inherit;
        }

        .electoraNavDescription {
          display: block;

          margin-top: 4px;

          font-size: 10px;

          line-height: 1.2;

          color: #9b958c;

          white-space: break-words;

          overflow: hidden;

          text-overflow: ellipsis;
        }

        .electoraNavItemActive
        .electoraNavDescription {
          color: #78899a;
        }

        /* =====================================================
           ARROW
        ===================================================== */

        .electoraNavArrow {
          flex-shrink: 0;

          margin-left: 8px;

          color: #c1bbb2;

          opacity: 0;

          transform: translateX(-3px);

          transition:
            opacity 0.16s ease,
            transform 0.16s ease;
        }

        .electoraNavItem:hover
        .electoraNavArrow {
          opacity: 1;

          transform: translateX(0);
        }

        .electoraNavItemActive
        .electoraNavArrow {
          opacity: 1;

          transform: translateX(0);

          color: #98743e;
        }

        /* =====================================================
           BOTTOM
        ===================================================== */

        .electoraSidebarBottom {
          flex-shrink: 0;

          padding: 0 16px 18px;
        }

        .electoraBottomDivider {
          height: 1px;

          background: #e8e5df;

          margin-bottom: 15px;
        }

        .electoraVersionInfo {
          display: flex;

          align-items: center;

          padding: 0 3px;
        }

        .electoraVersionIcon {
          width: 31px;
          height: 31px;

          flex-shrink: 0;

          display: flex;

          align-items: center;
          justify-content: center;

          background: #f4f2ee;

          border: 1px solid #e1ddd5;

          border-radius: 6px;

          color: #17395f;

          font-family:
            Georgia,
            "Times New Roman",
            serif;

          font-size: 14px;

          font-weight: 600;
        }

        .electoraVersionText {
          margin-left: 9px;

          display: flex;

          flex-direction: column;
        }

        .electoraVersionText span {
          color: #625d56;

          font-size: 10px;

          font-weight: 600;
        }

        .electoraVersionText small {
          margin-top: 2px;

          color: #aaa49a;

          font-size: 9px;
        }

        .electoraVersionNumber {
          margin-left: auto;

          color: #aaa49a;

          font-size: 9px;
        }

        /* =====================================================
           MOBILE BUTTON
        ===================================================== */

        .electoraMobileButton {
          display: none;

          position: fixed;

          top: 14px;
          left: 14px;

          width: 43px;
          height: 43px;

          z-index: 900;

          align-items: center;
          justify-content: center;

          border: 1px solid #dedbd4;

          border-radius: 7px;

          background: #ffffff;

          color: #17395f;

          cursor: pointer;

          box-shadow:
            0 3px 12px
            rgba(20, 25, 30, 0.08);
        }

        /* =====================================================
           OVERLAY
        ===================================================== */

        .electoraOverlay {
          display: none;

          position: fixed;

          inset: 0;

          z-index: 999;

          background:
            rgba(17, 24, 31, 0.38);

          backdrop-filter: blur(2px);
        }

        /* =====================================================
           TABLET
        ===================================================== */

        @media (max-width: 900px) {

          .electoraSidebar {
            width: 255px !important;
          }

        }

        /* =====================================================
           MOBILE
        ===================================================== */

        @media (max-width: 700px) {

          .electoraSidebar {
            width: 285px !important;

            transform:
              translateX(-100%) !important;

            box-shadow:
              8px 0 30px
              rgba(0, 0, 0, 0.12);

            transition:
              transform 0.25s ease;
          }

          .electoraSidebar.electoraSidebarOpen {
            transform:
              translateX(0) !important;
          }

          .electoraCloseButton {
            display: flex;
          }

          .electoraMobileButton {
            display: flex;
          }

          .electoraOverlay {
            display: block;
          }

          .electoraBrand {
            height: 76px;

            min-height: 76px;
          }

          .electoraNavigation {
            padding-top: 35px;
          }

        }

        /* =====================================================
           SMALL MOBILE
        ===================================================== */

        @media (max-width: 360px) {

          .electoraSidebar {
            width: 270px !important;
          }

          .electoraBrand {
            padding-left: 15px;
            padding-right: 15px;
          }

          .electoraSystemStatus {
            margin-left: 12px;
            margin-right: 12px;
          }

        }

      `}</style>
        </>
    );
}

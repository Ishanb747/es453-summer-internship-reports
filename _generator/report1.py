"""Report 1 – Comprehensive Training Report (ES-453, Project/Industrial track)."""
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm

import doclib as L
from data import STUDENT, ORG, WEEKS

SP = "/private/tmp/claude-501/-Users-ishan-Documents-New-project-Summer-Intern-report/76f2bf2a-0fb3-4818-a4c4-dc9469e7c347/scratchpad"
DIA, DL, SHOT, PAGES = f"{SP}/diagrams", f"{SP}/dl", f"{SP}/shots", f"{SP}/pages"
C = WD_ALIGN_PARAGRAPH.CENTER


LOGO = f"{SP}/logo/mait.jpg"


def cover(d, report_title):
    def line(parts, size, after=4, before=0):
        p = d.add_paragraph()
        p.alignment = C
        pf = p.paragraph_format
        pf.space_after, pf.space_before, pf.line_spacing = Pt(after), Pt(before), 1.15
        for text, bold, italic in parts:
            r = p.add_run(text); r.bold = bold; r.italic = italic; r.font.size = Pt(size)
        return p
    line([(report_title, True, False)], 22, after=10, before=24)
    line([("Subject Code: ES-453 (7th Semester)", False, False)], 12, after=14)
    line([(STUDENT["title"], True, False)], 14, after=18)
    line([("Submitted in partial fulfilment of the", False, False)], 12, after=0)
    line([("Requirements for the award of", False, False)], 12, after=6)
    line([("Degree of Bachelor of Technology in Information Technology", True, False)], 13, after=26)
    p = d.add_paragraph(); p.alignment = C
    p.paragraph_format.space_after = Pt(26)
    p.add_run().add_picture(LOGO, width=Cm(4.2))
    line([("Submitted By:", False, False)], 13, after=2)
    line([("Name of Student: ", False, False), (STUDENT["name"], True, True)], 13, after=2)
    line([("Enrollment No.: ", False, False), (f"{STUDENT['roll']} ({STUDENT['group'].replace('-', '')})", True, True)], 13, after=18)
    line([("Under the Guidance of:", False, False)], 13, after=2)
    line([("Faculty Guide: ", False, False), ("______________________", True, True)], 13, after=2)
    line([("Industry Mentor: ", False, False), (f"{ORG['mentor']}, {ORG['name']}", True, True)], 13, after=2)
    line([("Head of Department: ", False, False), ("Prof. (Dr.) Amita Goel", True, True)], 13, after=30)
    line([("SUBMITTED TO:", True, False)], 14, after=2)
    line([("Department of Information Technology", True, False)], 14, after=2)
    line([("MAHARAJA AGRASEN INSTITUTE OF TECHNOLOGY", False, False)], 13, after=2)
    line([("GGSIPU, DELHI", False, False)], 13, after=2)
    line([("Academic Session 2026–27", False, False)], 12, after=0)


REPORT1_INDEX = {
    "contents": [
        ("Certificate by Company", "CERTIFICATE BY COMPANY"), ("Declaration", "DECLARATION"),
        ("Acknowledgement", "ACKNOWLEDGEMENT"), ("About the Company", "ABOUT THE COMPANY"),
        ("List of Tables", "LIST OF TABLES"), ("List of Figures", "LIST OF FIGURES"),
        ("Abbreviations and Nomenclature", "ABBREVIATIONS AND NOMENCLATURE"),
        ("Chapter 1: Introduction", "CHAPTER 1"), ("Chapter 2: Tools & Technology Used", "CHAPTER 2"),
        ("Chapter 3: Technical Contents – Work Performed", "CHAPTER 3"), ("Chapter 4: Screenshots", "CHAPTER 4"),
        ("Chapter 5: Results and Discussions", "CHAPTER 5"), ("Chapter 6: Conclusions and Future Scope", "CHAPTER 6"),
        ("Weekly Jobs Summary", "WEEKLY JOBS SUMMARY"), ("Learning Outcome after Training", "LEARNING OUTCOME AFTER TRAINING"),
        ("References", "REFERENCES"),
    ],
    "annexure": [
        ("Appendix D – Weekly Logbooks", "APPENDIX D"), ("Appendix F – Industry Mentor Feedback Form", "APPENDIX F"),
        ("Appendix G – Student Feedback Form", "APPENDIX G"), ("Appendix H – Course Exit Survey", "APPENDIX H"),
        ("Appendix J – CO–PO/PSO Mapping", "APPENDIX J"), ("Appendix K – SDG Mapping", "APPENDIX K"),
    ],
}


def index_page(d, spec, pages, first=False):
    h = L.h1(d, "INDEX")
    L.no_outline(h)
    if first:
        h.paragraph_format.page_break_before = False
    for key, title in (("contents", "Table of Contents:"), ("annexure", "Annexure:")):
        if key not in spec:
            continue
        p = L.para(d, title, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, size=13, space_after=10)
        p.paragraph_format.space_before = Pt(10 if key == "contents" else 18)
        for i, (label, marker) in enumerate(spec[key], 1):
            L.index_line(d, i, label, (pages or {}).get(marker, "   "))


def certificate_page(d):
    L.h1(d, "CERTIFICATE BY COMPANY")
    L.para(d, "Internship certificate issued by PV Doctor Pte. Ltd. on its letterhead (4 September 2026).", italic=True, align=C, size=11)
    L.full_page_image(d, f"{PAGES}/cert-1.png", width_cm=15.0)


def declaration(d, heading="DECLARATION"):
    L.h1(d, heading)
    L.para(d, f"I hereby declare that the Training Report entitled \"{STUDENT['title']}\" is an authentic record of my own work as requirements of "
              f"{STUDENT['weeks']} weeks Training during the period from 5 May 2026 to 31 August 2026 for the award of degree of B.Tech. (Information Technology), GGSIPU, "
              f"under the guidance of {ORG['mentor']}, {ORG['mentor_desig']}, {ORG['name']}, Singapore.")
    L.para(d, "")
    for t in ("Signature of Student: ______________________________",
              f"Name of Student: {STUDENT['name']}",
              f"University Roll No.: {STUDENT['roll']}",
              f"Date: {STUDENT['sign_date']}"):
        L.para(d, t, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=2)
    L.para(d, "")
    L.para(d, "Certified that the above statement made by the student is correct to the best of our knowledge and belief.")
    L.para(d, "Examined by:", bold=True)
    L.para(d, "")
    L.para(d, "1. ______________________________          2. ______________________________", align=WD_ALIGN_PARAGRAPH.LEFT, space_after=0)
    L.para(d, "    Industry Mentor / Research Supervisor            Faculty Counsellor", align=WD_ALIGN_PARAGRAPH.LEFT)
    L.para(d, "")
    L.para(d, "Head of Department (Signature and Seal): ______________________________", align=WD_ALIGN_PARAGRAPH.LEFT)


def acknowledgement(d):
    L.h1(d, "ACKNOWLEDGEMENT")
    L.para(d, "I would like to express my sincere gratitude to PV Doctor Pte. Ltd., Singapore, for giving me the opportunity to work as an intern in its "
              "Data Engineering & Analytics team from May to August 2026. Working on a live product used by solar asset owners in more than twenty countries "
              "was the most valuable learning experience of my degree so far.")
    L.para(d, "I am deeply thankful to my industry mentor, Mr. Manav Gupta, Head of Data Engineering & Analytics, who guided my work throughout the internship, reviewed every pull request "
              "patiently and explained the reasoning behind each design decision. I also thank Dr. André Nobre, Managing Director, and Ms. Dannila Choong, Head of Finance & "
              "General Affairs, for their support with the internship formalities, and all members of the team who answered my questions about solar "
              "performance analysis and the portal codebase.")
    L.para(d, "I am grateful to Prof. (Dr.) Amita Goel, Head of the Department of Information Technology, Maharaja Agrasen Institute of Technology, "
              "Mr. Ajay Kumar Kaushik, Summer Training Coordinator, and my Faculty Counsellor for their guidance on the ES-453 training process and this report.")
    L.para(d, "Finally, I thank my family and friends for their constant encouragement.")
    L.para(d, "")
    L.para(d, STUDENT["name"], bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=0)
    L.para(d, f"University Roll No.: {STUDENT['roll']}", align=WD_ALIGN_PARAGRAPH.RIGHT)


def about_company(d):
    L.h1(d, "ABOUT THE COMPANY")
    L.para(d, "**PV Doctor Pte. Ltd.** is a Singapore-based Asset Performance Management (APM) provider for solar photovoltaic (PV) and battery energy storage (BESS) "
              "systems. The company was incorporated on 21 October 2024 (UEN 202443177E) and is headquartered at 18 Cross Street, #02-101, Singapore 048423. "
              "Its stated mission is to act as the \"trusted care-taker\" of a PV asset: it listens to the problems of an under-performing system, analyses the data "
              "to find root causes, and stays with the client until the system is rectified and back to optimum performance.")
    L.h2(d, "Services")
    L.bullets(d, [
        "**Smart O&M and Data Analytics** – continuous remote monitoring of PV systems through the Smart O&M Portal, daily performance reports, KPI-based alarms and analytics.",
        "**Preventive O&M** – scheduled inspections and maintenance planning driven by monitoring data.",
        "**Rectifications & Investigations** – root-cause analysis of under-performance and support until the fault is fixed.",
        "**Repowering Services** – upgrading ageing or damaged PV systems.",
        "**Audits & Performance Assessments** – independent technical audits of PV plants.",
        "**Technical Analytics** – custom studies such as degradation, soiling and yield analyses.",
    ])
    L.h2(d, "Scale and Growth")
    L.para(d, "According to the company's public website, PV Doctor manages more than 2.5 GWp of PV capacity across over 2,600 systems in about 25 countries "
              "(as of September 2026). Its clients range from residential owners to commercial and industrial rooftops, utility-scale solar farms, floating "
              "systems, educational and government institutions, and a large international pharmaceutical company. Figure 1 shows how quickly the managed capacity has grown "
              "since the company began operating in 2024.")
    L.figure(d, f"{DIA}/fig_growth.png", "Growth of PV capacity managed by PV Doctor (Q2 2024 – Q3 2026)")
    L.h2(d, "Organisation")
    L.para(d, "The company is led by its Managing Director, Dr. André Nobre. Interns and staff work in functional teams that include Business Development, "
              "Systems Engineering, IT & Data Services, Performance Engineering, Artificial Intelligence, and Data Engineering & Analytics (DEA). "
              "I was attached to the **Data Engineering & Analytics** team, which owns the data pipelines and the Smart O&M Portal "
              "(portal.pv.doctor), reporting to the Head of Data Engineering & Analytics. The team works remotely across Singapore, India and other countries "
              "and coordinates through Jira, code reviews and regular online meetings.")
    L.table(d, ["Particular", "Details"], [
        ["Company", "PV Doctor Pte. Ltd."],
        ["Registered address", "18 Cross Street, #02-101, 18 Cross, Singapore 048423"],
        ["Registration (UEN)", "202443177E, incorporated 21 October 2024"],
        ["Business", "Asset Performance Management for solar PV and BESS"],
        ["Main product", "Smart O&M Portal – portal.pv.doctor"],
        ["Website", "www.pv.doctor"],
        ["Department (intern)", "Data Engineering & Analytics"],
        ["Internship role", "Data Analytics Intern (scope of work: Web Development of the Monitoring Portal)"],
        ["Mode / Duration", "Remote; 5 May 2026 – 31 August 2026 (17 weeks)"],
    ], widths=[5, 11], title="Company profile at a glance")


def front_lists(d):
    L.no_outline(L.h1(d, "LIST OF TABLES"))
    L.toc(d, 'TOC \\h \\z \\c "Table"')
    L.no_outline(L.h1(d, "LIST OF FIGURES"))
    L.toc(d, 'TOC \\h \\z \\c "Figure"')
    L.h1(d, "ABBREVIATIONS AND NOMENCLATURE")
    L.table(d, ["Abbreviation / Symbol", "Meaning"], [
        ["AC", "Alternating Current"], ["API", "Application Programming Interface"], ["APM", "Asset Performance Management"],
        ["BESS", "Battery Energy Storage System"], ["CM / LT / LY", "Current Month / Lifetime / Last Year (period filters)"],
        ["CoV", "Coefficient of Variation"], ["CSS", "Cascading Style Sheets"], ["GHI", "Global Horizontal Irradiance"],
        ["G1 / G2 / G3", "Graph groups on the Inverter / MFM monitoring page"], ["INV", "Inverter"],
        ["KPI", "Key Performance Indicator"], ["kWh/kWp", "Specific yield – energy per installed peak kilowatt"],
        ["MFM", "Multi-Function Meter (energy meter)"], ["MWp / GWp", "Megawatt-peak / Gigawatt-peak (installed DC capacity)"],
        ["O&M", "Operations and Maintenance"], ["PNG", "Portable Network Graphics (image export)"], ["POA", "Plane of Array (irradiance on module plane)"],
        ["PR", "Performance Ratio"],
        ["PV", "Photovoltaic"], ["REST", "Representational State Transfer"], ["SDG", "Sustainable Development Goal"],
        ["SQL", "Structured Query Language"], ["STC", "Standard Test Conditions (1000 W/m², 25 °C)"], ["UI / UX", "User Interface / User Experience"],
        ["H", "In-plane irradiation over a period (kWh/m²)"], ["E", "AC energy delivered over a period (kWh)"],
        ["P<sub>0</sub>", "Installed DC capacity of the system (kWp)"], ["σ, μ", "Standard deviation and mean of a set of values"],
    ], widths=[4.5, 11.5], title="Abbreviations and symbols", size=10)


# ------------------------------------------------------------------ chapters

def chapter1(d):
    L.h1(d, "CHAPTER 1\nINTRODUCTION")
    L.h2(d, "1.1 Background")
    L.para(d, "Solar photovoltaic capacity is growing faster than any other source of electricity. India alone has more than 25 GW of rooftop solar and several "
              "lakh new household systems every year, and South-East Asian countries are adding utility-scale and commercial rooftop plants at a similar pace. "
              "Once a plant is built, however, its output depends on how well it is operated. Faults such as inverter trips, string failures, soiling, shading and "
              "communication loss can silently reduce generation for weeks if nobody is watching the data.")
    L.para(d, "Asset Performance Management companies such as PV Doctor solve this problem with continuous remote monitoring. Data from inverters, energy meters "
              "(MFMs) and weather stations is collected every few minutes, stored in a central database and converted into Key Performance Indicators (KPIs) such as "
              "specific yield, Performance Ratio (PR) and Coefficient of Variation (CoV). Engineers and asset owners then use a web portal to see which plants and which "
              "devices are under-performing and why.")
    L.para(d, "The usefulness of such a portal depends heavily on its visualisations. A chart with wrong units, a date filter that silently drops a month, or a "
              "tooltip that is cut off at the edge of the screen can mislead an engineer and delay a repair. My internship was focused on exactly this layer: building and "
              "refining the interactive charts, filters and data views of PV Doctor's Smart O&M Portal, and the backend APIs that feed them.")
    L.h2(d, "1.2 About the Training")
    L.table(d, ["Item", "Details"], [
        ["Organisation", "PV Doctor Pte. Ltd., Singapore"],
        ["Team", "Data Engineering & Analytics"],
        ["Designation", "Data Analytics Intern (Web Development of the Monitoring Portal)"],
        ["Duration", "5 May 2026 – 31 August 2026 (17 weeks)"],
        ["Mode", "Remote (work from home, India)"],
        ["Working hours", "9 am – 6 pm, Monday to Friday until 31 July; about 4 hours per day in August"],
        ["Stipend", "INR 20,000 per month"],
        ["Project", STUDENT["title"]],
    ], widths=[4, 12], title="Internship details")
    L.para(d, "The internship agreement (Appendix A of the agreement) defined the scope of work as: designing and developing PV Doctor's Monitoring Portal, including "
              "dashboards and core user interface components; building responsive front-end applications in React with data-visualisation libraries; integrating the "
              "front end with backend services powered by FastAPI; following clean-architecture principles with attention to performance, scalability and reliability; "
              "and taking end-to-end ownership of assigned features from requirement analysis to delivery. The original term was 5 May to 31 July 2026; it was extended "
              "to 31 August 2026 on a part-time basis, as recorded in the company's internship certificate.")
    L.h2(d, "1.3 Objectives")
    L.bullets(d, [
        "Understand the solar PV monitoring domain – the meaning of yield, irradiation, PR, CoV and corridor analysis – well enough to present them correctly.",
        "Build and improve interactive charts for plant-level and device-level monitoring (Daily Generation, Irradiation vs Yield, PR Evolution, Yield–PR scatter, CoV, Heatmap, Stacked Bar, Spatial Heatmap).",
        "Make period and date-range filtering correct and consistent across all modules.",
        "Reduce duplicated code by refactoring chart implementations into shared components.",
        "Extend the FastAPI backend for System Information, Raw Data, reference sensors and satellite (SolarGIS) data, with attention to query performance.",
        "Follow professional software practice: Jira tasks, feature branches, pull requests, code review, staging verification and automated tests.",
    ])
    L.h2(d, "1.4 Scope of this Report")
    L.para(d, "Company source code, client names, credentials and datasets are confidential under the internship agreement. This report therefore describes the "
              "work at the level of features, design decisions and algorithms, uses generic descriptions instead of client site names, and uses the public material "
              "from portal.pv.doctor and the company website for illustrations.")
    L.h2(d, "1.5 Organisation of the Report")
    L.para(d, "Chapter 2 describes the tools and technologies used. Chapter 3 presents the technical work module by module. Chapter 4 contains screenshots and "
              "visuals of the portal. Chapter 5 discusses the results and challenges, and Chapter 6 concludes with future scope. The weekly jobs summary, learning "
              "outcomes, references and the prescribed appendices (weekly logbooks, feedback forms, exit survey, CO–PO/PSO and SDG mapping) follow.")


def chapter2(d):
    L.h1(d, "CHAPTER 2\nTOOLS & TECHNOLOGY USED")
    L.para(d, "The Smart O&M Portal is a full-stack web application. The front end is a Next.js application written in TypeScript, the back end is a Python FastAPI "
              "service, and data is stored in a SQL Server database. Both services run behind Nginx on a Linux server, with separate production and staging instances "
              "managed by systemd. Figure 2 shows the overall architecture.")
    L.figure(d, f"{DIA}/fig_arch.png", "High-level architecture of the Smart O&M Portal")
    L.h2(d, "2.1 Front-end Technologies")
    L.table(d, ["Technology", "Purpose in the project"], [
        ["Next.js (App Router)", "React framework used for routing (portfolio-overview, monitoring, reports, admin, profile), layouts and production builds."],
        ["React", "Component model for charts, filters, tables and panels; hooks and context for user, client and preference state."],
        ["TypeScript", "Static typing of API responses, chart data and component props; catches errors before runtime."],
        ["Chart.js 4", "Main charting library for bar, line and scatter charts (generation, yield, PR, CoV, stacked bars)."],
        ["chartjs-plugin-zoom", "Zoom and pan behaviour, restricted to the X-axis where needed."],
        ["chartjs-adapter-date-fns / date-fns", "Time axes and date arithmetic for period filters (CM, LT, LY, custom ranges)."],
        ["Plotly.js (react-plotly.js)", "Heatmaps and interactive visualisations that need colour scales and hover detail."],
        ["Google Maps JavaScript API", "World map in Portfolio Overview and satellite layer for the Spatial Heatmap."],
        ["SheetJS (xlsx)", "Client-side Excel export of tables and chart data."],
        ["CSS Modules", "Component-scoped styling, responsive breakpoints and theming."],
        ["lucide-react, lottie-react", "Icons and lightweight animations."],
    ], widths=[4.8, 11.2], title="Front-end technologies")
    L.h2(d, "2.2 Back-end Technologies")
    L.table(d, ["Technology", "Purpose in the project"], [
        ["Python 3 + FastAPI", "REST API with route modules: auth, monitoring, portfolio_overview, files, stats, preferences, contact, health, admin."],
        ["Uvicorn", "ASGI server running the FastAPI application."],
        ["pyodbc + SQL Server", "Database access for site metadata, device time series, daily summaries, alarms and users."],
        ["slowapi", "Rate limiting of API endpoints by client address."],
        ["itsdangerous, python-dotenv", "Signed tokens and environment-based configuration (no secrets in code)."],
        ["SolarGIS data", "Satellite-derived irradiance used as a reference when on-site sensors are missing or unreliable."],
        ["Automated tests", "API tests for auth, files, contact, health, preference and stats modules (added in week 17)."],
    ], widths=[4.8, 11.2], title="Back-end technologies")
    L.h2(d, "2.3 Deployment and Collaboration Tools")
    L.table(d, ["Tool", "Use"], [
        ["Git & GitHub", "Version control, feature branches and pull requests with code review."],
        ["Jira (Atlassian)", "Task tracking for feature requests and review comments."],
        ["Nginx", "Reverse proxy for the web and API services of portal.pv.doctor."],
        ["systemd", "Service units for production and staging (web and API)."],
        ["VS Code", "Development environment."],
        ["Chrome DevTools", "Debugging layout, network calls and performance in the browser."],
        ["Microsoft Excel", "Verifying exported data and maintaining the intern task tracker."],
        ["Google Meet / chat", "Daily communication with the remote team."],
    ], widths=[4.8, 11.2], title="Deployment and collaboration tools")
    L.h2(d, "2.4 Domain Concepts")
    L.para(d, "Correct visualisation requires correct domain understanding. The following KPIs, defined in the IEC 61724-1 standard for PV performance monitoring, "
              "appear throughout the portal:")
    L.bullets(d, [
        "**Specific (final) yield**:  Y<sub>f</sub> = E / P<sub>0</sub>  [kWh/kWp] – energy produced per kilowatt-peak installed; allows plants of different sizes to be compared.",
        "**Reference yield**:  Y<sub>r</sub> = H / G<sub>STC</sub>  [h] – in-plane irradiation divided by the STC irradiance of 1 kW/m²; the number of 'peak sun hours'.",
        "**Performance Ratio**:  PR = Y<sub>f</sub> / Y<sub>r</sub> × 100 % – the share of the theoretically available energy actually delivered; typically 75–85 % for a healthy plant.",
        "**Coefficient of Variation**:  CoV = σ / μ × 100 % – computed across inverters or meters of a site; a high CoV means some devices are performing differently from their peers.",
        "**Corridor** – an expected band for PR or yield derived from historical behaviour; points outside the corridor indicate abnormal performance and can raise alarms.",
        "**Irradiation vs irradiance** – irradiance is instantaneous power (W/m²); irradiation is energy over time (kWh/m²). The portal labels were corrected to use the right term.",
        "**CO₂ avoided** – energy generated multiplied by a grid emission factor; shown per site in Portfolio Overview.",
    ])
    L.figure(d, f"{DIA}/fig_workflow.png", "Development workflow followed for every task")
    L.para(d, "Every task followed the workflow in Figure 3: requirement from Jira or a review comment, analysis of the existing component and API, implementation "
              "on a feature branch, testing across date ranges, sites and screen sizes, a pull request reviewed by the team, and verification on staging before release.")


def chapter3(d):
    L.h1(d, "CHAPTER 3\nTECHNICAL CONTENTS – WORK PERFORMED")
    L.para(d, "This chapter describes the work module by module. Figure 4 shows the structure of the portal and where each module sits. For every module, "
              "the problem, the approach and the implementation are described.")
    L.figure(d, f"{DIA}/fig_modules.png", "Module structure of the Smart O&M Portal")
    L.figure(d, f"{DIA}/fig_timeline.png", "Timeline of work streams across the 17 weeks")

    L.h2(d, "3.1 Onboarding and Domain Study (Week 1)")
    L.para(d, "The first days were spent getting access to Jira and the code repository, setting up the frontend (Next.js development server) and backend "
              "(FastAPI with Uvicorn) locally, and reading through the existing modules. In parallel I studied the PV KPIs listed in Section 2.4 and the IEA PVPS "
              "'Trends in Photovoltaic Applications' material shared within the team, so that I could judge whether a chart was showing the right quantity in the right unit.")

    L.h2(d, "3.2 Plant Monitoring Graphs (Weeks 1–4)")
    L.para(d, "The Plant Monitoring page shows site-level performance through four main charts: Daily Generation (bar), Irradiation vs Yield (combined bar/line), "
              "PR Evolution (scatter with moving average and corridor) and a Yield–PR scatter. The first month was spent bringing these charts to production quality "
              "through several rounds of review.")
    L.h3(d, "3.2.1 Chart controls and menus")
    L.bullets(d, [
        "Moved download options into a vertical three-dot menu to the right of the chart toggles; downloads follow a fixed file-naming convention.",
        "Limited zoom to levels 0–9 and stacked the zoom-in, zoom-out and reset buttons vertically with semi-transparent styling, later made more opaque after feedback.",
        "Added a buffer above bars when the average line reaches or exceeds the y-axis limit, and blank space on the right of bar graphs, so labels are never clipped.",
        "Removed redundant 'period' and 'type' labels from chart headers and moved units into the column titles.",
    ])
    L.h3(d, "3.2.2 Tooltips, units and precision")
    L.para(d, "Tooltips were redesigned: keys left-aligned in white, values right-aligned in the accent colour, with the date shown in a consistent format. "
              "Units were corrected – yield in kWh/kWp, generation in kWh, irradiation in kWh/m² – and 'irradiance' was renamed 'irradiation' where energy over "
              "time is shown. Precision rules were fixed: two decimals for yield and irradiation, one decimal for daily generation, thousand separators for large values.")
    L.h3(d, "3.2.3 Multi-year analysis, regression and playback")
    L.bullets(d, [
        "Added a **year filter**: each year is drawn in a different shade of blue; deselected years become transparent instead of disappearing, so the axis does not jump. In the PR graph, past years are grey.",
        "Drew the **regression line** in red and limited the PR y-axis to 100 % with blank space above.",
        "Built a **playback** feature that animates the yield chart over time, with speed selection from a dropdown, chevron controls and a date indicator.",
        "Added a **screenshot effect** – a camera-shutter style flash on the chart border when the chart is captured as an image.",
    ])
    L.h3(d, "3.2.4 Corridor and layering fixes")
    L.para(d, "In PR Evolution, all points inside the corridor are coloured while points outside are shown in grey in the irradiation chart; the colour bar was toned "
              "down without changing its structure; the moving-average line was given a higher z-order and greater width so it stays visible above the scatter; and a bug "
              "where the PR Evolution chart was drawn beneath the full-width Plant Monitoring section was fixed. Daily Generation was changed to plot generation (kWh) instead of yield, "
              "with Current Year as the default and a working Lifetime (LT) view.")

    L.h2(d, "3.3 Portfolio Overview (Weeks 2–4, 7)")
    L.para(d, "Portfolio Overview is the landing page after login: a world map with every site coloured by status, KPI cards, a sites table, an alert panel and "
              "top/bottom performers. My work on this page included:")
    L.bullets(d, [
        "When several sites are selected, the map bounds are recomputed so that all selected points are visible; the default view shows the whole portfolio.",
        "Fixed tooltip cut-off and out-of-bounds behaviour on the map, reduced the search box size and made the Reset button bolder.",
        "Export of the map overview to **Excel** (with separate latitude and longitude columns and the system code) and to **PNG** (map, buffer and title).",
        "Made **Reset** clear the component state as well as the selected filters, so map filters, table and KPIs stay synchronised.",
        "Fixed decimal precision and rounding in the Portfolio Overview metrics and in the Spatial Heatmap tooltips.",
        "Added rules to hide sensitive sections for **demo accounts**, which are used to show the portal to prospective clients.",
        "Improved the responsive layout of the Overview KPI cards for small screens by adjusting breakpoints (week 9).",
    ])

    L.h2(d, "3.4 Period Filter and Date-Range Logic (Weeks 5–8)")
    L.para(d, "Many bugs reported by users came from inconsistent handling of time periods across charts. For example, selecting 'Current Month + 2026' in Daily Generation "
              "displayed only June, and the Hindsight view showed only the last month of the selected year. I corrected the period-selection logic and then defined explicit "
              "rules, agreed with the team, so that every chart behaves the same way. Table 7 summarises them.")
    L.table(d, ["Selection", "Rule implemented"], [
        ["Today / Yesterday / single date", "Show intraday data; date label format 'D' (day); Heatmap behaviour matched to CoV."],
        ["Current Month (CM)", "Daily values of the current month; label 'CM'."],
        ["Last Year (LY) / single year", "Show the full range of that year only; no extra sub-filters."],
        ["Lifetime (LT)", "Whole history of the site; label 'LT' instead of 'All years'."],
        ["Custom range (two dates)", "Default to LT granularity with a dropdown to change it."],
        ["Custom range of a single day", "Default to the last 30 days (30D) with a dropdown."],
        ["Custom range < 365 days spanning two years", "Monthly aggregation with Month + Year labels."],
        ["Only one option available", "Hide the filter instead of showing a dropdown with a single entry."],
    ], widths=[5.5, 10.5], title="Period and date-range rules")
    L.para(d, "I also restricted PR Evolution zoom to the X-axis, fixed the CoV monthly X-axis, and later (week 14) synchronised date handling across PR Evolution, Raw Data and Spatial "
              "Heatmap so that the same period means the same dates everywhere. In week 8 the System Information view was connected to the Period Filter, which required changes "
              "in the backend monitoring route to support custom access levels; this was tested end-to-end across multiple sites and user roles.")

    L.h2(d, "3.5 Inverter and Meter (MFM) Monitoring (Weeks 3, 6–8)")
    L.para(d, "Device-level monitoring compares inverters and energy meters of the same site. Charts are organised in groups G1 (generation bars), G2 (stacked bar) "
              "and G3 (PR/CoV/heatmap). Work on this page included:")
    L.bullets(d, [
        "Making the Inverter page consistent with Plant Monitoring: zoom instead of clipping, separate (not joined) grids, the same subtle grid style and axis formatting.",
        "Inverter uniformity chart with legends outside the plot and the green / yellow / red thresholds always visible.",
        "**Multi-select filters** for inverters and MFMs with a Reset action, auto-width based on selected text, and a compact 'All MFMs' label; G1/G2/G3 attached to the filters.",
        "Heatmap revisions: heatmap aligned with its colour bar, 'Legend' renamed 'Thresholds', legend split into two parts, consistent fonts.",
        "CoV chart: legend height equal to chart height, point size adapting to the Year/Day toggle, divider lines removed.",
        "Yield–PR graph: a Year/Day toggle, corrected Y-axis, higher z-order for selected items, and a fix for marker symbols (+, ×, *) that disappeared in some conditions.",
    ])

    L.h2(d, "3.6 Refactoring Chart Components (Week 8)")
    L.para(d, "The Inverter and Meter pages had separate copies of the CoV chart, the Heatmap and the G2 Stacked Bar Chart. Every fix had to be made twice and the two copies "
              "had started to drift apart. I refactored the three chart types into shared components that take the entity type (inverter or meter) and data as configuration, "
              "re-implemented Heatmap Playback on top of the shared component, tested all scenarios and raised the pull request. Figure 6 shows the change.")
    L.figure(d, f"{DIA}/fig_refactor.png", "Refactoring duplicated chart implementations into shared components")

    L.h2(d, "3.7 System Information, Admin and Alarms (Weeks 9–13)")
    L.para(d, "During July I worked on the operational views of the portal. The **System Information** view presents per-site time series – on-site irradiance sensor, "
              "clear-sky GHI, meter and inverter AC power, and export-meter energy – with consistent column names, units and Excel export. The **Admin** user-management view lists "
              "users with their role, client, status, pending-setup flag and last login, and supports an Excel snapshot of the user list. The **alarm views** show KPI-based alarms "
              "such as 'Full Site PR: Corridor Deviation' with severity, category, description and a recommended action, and an Alarm History export. Because these views contain "
              "personal and client data, exports were tested only with internal accounts and no data from them is reproduced in this report.")

    L.h2(d, "3.8 Spatial Heatmap – PV Healthscan (Weeks 13–16)")
    L.para(d, "The Spatial Heatmap (marketed as PV Healthscan) overlays yield deviation of each inverter or string block on the satellite image of the plant, so that "
              "under-performing areas are visible at a glance and cleaning or maintenance can be prioritised. I completed the interactive version of this view:")
    L.bullets(d, [
        "Responsive map container sized to the screen, consistent spacing between title, colour bar and map, and a pan icon on the graph.",
        "Conditional zoom rules so that sites of very different sizes are framed correctly; fixed an unwanted zoom when switching to the deviation view on some sites.",
        "Scrollable ranking list ordered from best to worst, with sites or devices that have no data still listed as N/A.",
        "Legend renaming logic with the full name available in tooltips; deviation shown with one decimal; refined colour scaling and statistics.",
        "Overlays and labels hidden on demand to see the underlying satellite image.",
    ])

    L.h2(d, "3.9 PR Evolution Corridor and Historical Averages (Weeks 14–15)")
    L.para(d, "I fixed the Net PR corridor calculation and updated the PR Evolution chart to show the corridor and historical averages correctly. On the device-level PR graph, "
              "the lifetime average line is also drawn when multiple inverters or meters are selected, with its value in the legend. Chart styling (tick boxes, legend font size "
              "and weight, dropdown sizing) was aligned with the CoV chart so that all monitoring views look the same.")

    L.h2(d, "3.10 Raw Data, Reference Sensors and SolarGIS (Weeks 14–16)")
    L.para(d, "The Raw Data view lets engineers download aligned time series for any combination of components. This part of the work was mostly backend:")
    L.bullets(d, [
        "**AC Power / AC Energy exclusivity** – a series is requested either as power or as energy, avoiding mixed and misleading columns.",
        "**Mixed-resolution granularity detection** – when components log at different intervals (e.g. 5-minute and 15-minute), the resolution is detected and the series are aligned on a common time axis.",
        "**Reference sensor support** – each component can be tagged to a reference irradiance sensor (stored as a dictionary keyed by component id); the result carries a ref_sensor column. If the tag is not the component's own sensor, only the ground-type sensor is used.",
        "**Batched external sensor resolution** – some sites use a sensor installed at a neighbouring site. Instead of one query per sensor, internal and external sensors are grouped and resolved in batches, which also fixed errors on multi-site sensor setups.",
        "**SolarGIS** – dynamic granularity for satellite irradiance and support for plants with modules in multiple orientations.",
        "**Parallel queries and caching** – meter, inverter, sensor and SolarGIS series are fetched concurrently and cached, reducing response time for large date ranges.",
    ])
    L.figure(d, f"{DIA}/fig_refsensor.png", "Raw-data request flow with batched reference-sensor resolution")

    L.h2(d, "3.11 Backend Refactoring and Automated Tests (Week 17)")
    L.para(d, "My pull request for the monitoring routes had grown to more than 700 lines. Following review feedback, I refactored it to 289 lines – about 59 % smaller – "
              "by removing duplicated query construction, extracting helper functions and keeping only the functionality that was required. I also added automated API tests for the "
              "auth and files routes and test files covering the contact, health, preference and stats modules, so that future changes can be verified automatically.")

    L.h2(d, "3.12 Support Widget Requirements (Week 17)")
    L.para(d, "At the end of the internship I analysed the requirements for an in-portal support feature: replacing the top-bar support icon with a floating bubble at the "
              "bottom right, a pop-up that routes solar-performance questions to the Performance Engineering team and portal questions to Smart O&M IT, separate WhatsApp and call "
              "buttons, and redirecting users who click a locked feature to the support page with an explanatory hover message. These notes were handed over to the team.")


def chapter4(d):
    L.h1(d, "CHAPTER 4\nSCREENSHOTS")
    L.para(d, "The screenshots below were taken from the public demo account of portal.pv.doctor (October 2026). The demo uses anonymised sites "
              "(\"Site Code 1–7\"), so no client data is shown. Each figure is a module I built or worked on during the internship; the relevant "
              "section of Chapter 3 is given in brackets.")
    CARDS = f"{SP}/cards"
    shots = [
        (f"{SHOT}/01_login.png", "Smart O&M Portal – login page with demo access", 15.5,
         "The public entry point of the portal. The Demo button opens a read-only demo account; demo accounts hide sensitive sections (Section 3.3)."),
        (f"{CARDS}/c_portfolio.jpg", "Portfolio Overview – KPI cards, portfolio map, system alerts, top/bottom performers and portfolio summary", 15.5,
         "Map auto-fit for selected systems, Reset that clears all state, Yield/PR toggle, decimal formatting and the Excel/PNG export menu (Section 3.3)."),
        (f"{CARDS}/c_daily_gen.jpg", "Plant Monitoring – Daily Generation with average line and highest/lowest markers", 15.5,
         "Generation plotted in kWh, Daily / Monthly / Lifetime toggles, buffer above the tallest bar, the dashed average line, and the three-dot download menu (Section 3.2)."),
        (f"{CARDS}/c_irr_yield.jpg", "Plant Monitoring – Irradiation vs Yield scatter with year-wise colouring and fit line", 15.5,
         "The x-axis is limited to 9, units are kWh/m² and kWh/kWp, and the Lifetime/Fit/Yield/PR controls and the star marking the latest day come from the Week 1–2 revisions (Section 3.2.3)."),
        (f"{CARDS}/c_pr_evo.jpg", "Performance Ratio Evolution – corridor (±10 %), 30-day moving average and lifetime average", 15.5,
         "Points inside and outside the corridor are coloured differently, the moving average is drawn above the scatter, 'Data Available = x days' is shown, and zoom is restricted to the X-axis (Sections 3.2.4, 3.4 and 3.9)."),
        (f"{CARDS}/c_inv_daily_gen.jpg", "Inverter monitoring – stacked Daily Generation per inverter (G2) with inverter multi-select", 15.5,
         "The shared Stacked Bar component, refactored from two copies, with the multi-select Inverters filter (Sections 3.5 and 3.6)."),
        (f"{CARDS}/c_inv_heatmap.jpg", "Inverter Performance Heatmap – daily yield per inverter over 30 days", 15.5,
         "The shared Heatmap component aligned with its colour bar and threshold scale (Sections 3.5 and 3.6)."),
        (f"{CARDS}/c_meter_heatmap.jpg", "Meters & Zones – Meter Yield Heatmap built from the same shared component", 15.5,
         "The same Heatmap component reused for energy meters (MFMs), which keeps the two pages consistent (Section 3.6)."),
        (f"{CARDS}/c_specific_yield.jpg", "Specific Yield per inverter – Lifetime / 365D / 90D views with base-inverter selection", 15.5,
         "Period toggles and filters follow the period rules in Table 7 (Section 3.4)."),
        (f"{CARDS}/c_spatial.jpg", "PV Spatial Heatmap (PV Healthscan) – inverter yield on the satellite layout with best-to-worst ranking", 15.5,
         "The interactive Spatial Heatmap: yield statistics (average, standard deviation, CoV), best-to-worst list, colour bar with marker, and the recentre control (Section 3.8)."),
        (f"{CARDS}/c_system_info.jpg", "System Information – intraday AC power against clear-sky profile", 15.5,
         "The System Information time series connected to the Period Filter (Sections 3.4 and 3.7)."),
        (f"{CARDS}/c_weather.jpg", "Weather Trends – comparison of ground irradiance sensors", 15.5,
         "Sensor-against-sensor comparison, related to the reference-sensor work in Raw Data (Section 3.10)."),
    ]
    for path, cap, w, text in shots:
        L.figure(d, path, cap, width_cm=w)
        d.paragraphs[-1].paragraph_format.keep_with_next = True
        p = L.para(d, text, size=11)
        p.paragraph_format.space_after = Pt(14)


def chapter5(d):
    L.h1(d, "CHAPTER 5\nRESULTS AND DISCUSSIONS")
    L.h2(d, "5.1 Results")
    L.para(d, "The work was delivered as a series of reviewed pull requests that were verified on staging and released to the production portal. Table 8 summarises "
              "the main outcomes.")
    L.table(d, ["Area", "Outcome"], [
        ["Plant Monitoring", "Four core charts brought to production quality: correct units and precision, year filters, regression, corridor colouring, playback and screenshot export."],
        ["Portfolio Overview", "Map auto-fit, reliable Reset, Excel/PNG export, demo-account visibility rules and responsive KPI cards."],
        ["Period filter", "One documented set of rules (Table 7) applied across charts; fixed month/year visibility and aggregation bugs."],
        ["Inverter / MFM", "Multi-select filters with reset; consistent styling with Plant Monitoring; improved heatmap, CoV and Yield–PR charts."],
        ["Code quality", "CoV, Heatmap and Stacked Bar charts consolidated from two copies each into shared components."],
        ["Spatial Heatmap", "Interactive PV Healthscan view completed with ranking list, zoom rules and refined scaling."],
        ["Raw Data backend", "Reference-sensor support, batched external sensor resolution, SolarGIS granularity and orientations, parallel queries with caching."],
        ["Backend refactor", "Monitoring pull request reduced from 700+ to 289 lines (≈59 % smaller) with the same required behaviour."],
        ["Testing", "Automated API tests added for auth and files plus contact, health, preference and stats modules."],
    ], widths=[4, 12], title="Summary of results")
    L.h2(d, "5.2 Discussion")
    L.para(d, "**Correctness before decoration.** The most important changes were often small: a wrong unit, an off-by-one in a date range, or a decimal rounding issue. "
              "Because O&M engineers act on these numbers, fixing them had more value than adding new visual features. Writing the period rules down as a table (Table 7) "
              "turned a stream of individual bug reports into one consistent behaviour.")
    L.para(d, "**Iterative review.** Most features went through two to four review rounds. Early in the internship this felt slow, but it produced a consistent "
              "product and taught me to anticipate questions – such as how a chart behaves for a single day, a single year or a site with only one inverter – before "
              "submitting work.")
    L.para(d, "**Reducing duplication pays back quickly.** After the chart refactor, later fixes (for example to legends and scrollbars in August) were made once instead of twice. "
              "Similarly, shrinking the backend pull request made it reviewable and easier to maintain.")
    L.para(d, "**Performance on real data.** Sites with many devices and long histories produce large queries. Batching sensor resolution, running independent queries in "
              "parallel and caching results were necessary to keep the Raw Data view responsive.")
    L.h2(d, "5.3 Challenges Faced")
    L.table(d, ["Challenge", "How it was handled"], [
        ["Learning a new domain (PV performance) quickly", "Studied IEC 61724 KPIs and asked the team to explain each chart's purpose before changing it."],
        ["Large existing codebase", "Read the existing component and its API before each change; kept changes small and focused."],
        ["Date and time-zone edge cases", "Built a test matrix of period selections (Today, CM, LY, LT, custom ranges) and checked each chart against it."],
        ["Remote collaboration across time zones", "Written updates in the task tracker, clear pull-request descriptions and screenshots for reviewers."],
        ["Site-specific data issues", "Reproduced using the affected site's data, then generalised the fix (e.g. multi-site reference sensors)."],
        ["Balancing internship with semester start (August)", "Moved to a fixed part-time schedule of about 4 hours per day."],
    ], widths=[6, 10], title="Challenges and how they were handled")


def chapter6(d):
    L.h1(d, "CHAPTER 6\nCONCLUSIONS AND FUTURE SCOPE")
    L.h2(d, "6.1 Conclusion")
    L.para(d, "During the 17-week internship at PV Doctor Pte. Ltd. I contributed to a production web application that helps solar asset owners and O&M engineers monitor "
              "more than 2.5 GWp of PV capacity. I built and refined the portal's main visualisations, made date and period handling consistent, refactored duplicated chart code "
              "into shared components, completed the interactive Spatial Heatmap, extended the FastAPI backend for raw data, reference sensors and satellite irradiance, and added "
              "automated tests.")
    L.para(d, "The internship met the objectives set in Chapter 1. Beyond the technical skills – React, Next.js, TypeScript, Chart.js, Plotly, FastAPI and SQL – I learned how "
              "professional teams work: tasks in Jira, small reviewed pull requests, staging before production, and careful handling of confidential client data. I also gained an "
              "understanding of how good software directly increases clean-energy output by helping engineers find faults sooner.")
    L.h2(d, "6.2 Future Scope")
    L.bullets(d, [
        "**Automated anomaly detection** – use machine-learning models on PR and CoV time series to detect soiling, shading and inverter faults earlier than fixed corridors.",
        "**Root-cause suggestions** – combine alarms with weather and device data to suggest the likely cause and recommended action automatically.",
        "**Front-end test coverage** – add component and end-to-end tests for chart and filter behaviour so that the period rules are verified on every change.",
        "**Support widget** – implement the in-portal support bubble and routing analysed in week 17.",
        "**Mobile experience and notifications** – a mobile-first view and WhatsApp/e-mail alerts for owners of small systems who do not use dashboards.",
        "**Performance** – pre-aggregated summaries and server-side caching for very large portfolios.",
    ])


def weekly_summary(d):
    L.h1(d, "WEEKLY JOBS SUMMARY")
    L.para(d, "Weeks 1–13 were full-time; weeks 14–17 were part-time (about 4 hours per day). Detailed weekly logbooks in the prescribed format are in Appendix D.")
    rows = []
    for w in WEEKS:
        rows.append([str(w["no"]), short_range(w["from"], w["to"]), w["activities"]])
    L.table(d, ["Wk", "Dates", "Jobs / tasks performed"], rows, widths=[1.0, 2.9, 12.1], title="Weekly jobs summary", size=10)


def short_range(a, b):
    da, ma, ya = a.split(); db, mb, yb = b.split()
    if ma == mb:
        return f"{da}–{db} {ma[:3]} {ya}"
    return f"{da} {ma[:3]} – {db} {mb[:3]} {yb}"


def learning_outcomes(d):
    L.h1(d, "LEARNING OUTCOME AFTER TRAINING")
    L.h2(d, "Technical Learning")
    L.bullets(d, [
        "Building production-quality React components in Next.js with TypeScript, including state management with hooks and context.",
        "Data visualisation with Chart.js and Plotly.js: scales, time axes, zoom and pan, tooltips, layering, colour for categories and time, and export to image and Excel.",
        "Designing consistent date-range and aggregation logic and testing it systematically.",
        "Developing REST APIs in FastAPI, writing efficient SQL access with batching, parallel queries and caching, and writing automated API tests.",
        "Refactoring for maintainability: removing duplication, extracting shared components and keeping pull requests small.",
        "Solar PV performance analysis: yield, irradiation, PR, CoV, corridors and the role of satellite irradiance.",
    ])
    L.h2(d, "Professional Learning")
    L.bullets(d, [
        "Working in a remote, international team with written communication, task trackers and code review.",
        "Taking ownership of a feature from requirement to release, including staging verification.",
        "Handling confidential client data responsibly.",
        "Managing time between internship work and academic commitments.",
    ])
    L.h2(d, "Mapping to Course Outcomes (ES-453)")
    L.table(d, ["CO", "Course outcome", "How it was achieved"], [
        ["CO1", "Analyze and understand real-world IT problems through internship tasks.", "Analysed user-reported issues in charts, filters and data queries of a live monitoring portal."],
        ["CO2", "Design or assist in implementing solutions using specific technologies.", "Designed and implemented chart components, period rules, raw-data APIs and refactors (Web Development, data engineering)."],
        ["CO3", "Use tools, platforms or frameworks relevant to the domain efficiently.", "Next.js, React, TypeScript, Chart.js, Plotly, FastAPI, SQL Server, Git, Jira, Nginx."],
        ["CO4", "Prepare comprehensive technical reports and presentations.", "Weekly task tracker entries, pull-request descriptions, this report and the viva presentation."],
    ], widths=[1.3, 6.2, 8.5], title="Mapping of internship work to ES-453 course outcomes")


def references(d):
    L.h1(d, "REFERENCES")
    refs = [
        "PV Doctor Pte. Ltd., company website, https://www.pv.doctor (accessed October 2026).",
        "PV Doctor Smart O&M Portal, public login and demo pages, https://portal.pv.doctor (accessed October 2026).",
        "IEC 61724-1:2021, Photovoltaic system performance – Part 1: Monitoring, International Electrotechnical Commission, 2021.",
        "T. Dierauf, A. Growitz, S. Kurtz, J. L. B. Cruz, E. Riley and C. Hansen, \"Weather-Corrected Performance Ratio,\" NREL Technical Report NREL/TP-5200-57991, 2013.",
        "IEA PVPS, Trends in Photovoltaic Applications, International Energy Agency Photovoltaic Power Systems Programme (latest edition).",
        "Vercel, Next.js Documentation, https://nextjs.org/docs.",
        "Meta Open Source, React Documentation, https://react.dev.",
        "Chart.js Contributors, Chart.js Documentation (v4) and chartjs-plugin-zoom, https://www.chartjs.org/docs.",
        "Plotly Technologies Inc., Plotly JavaScript Open Source Graphing Library, https://plotly.com/javascript.",
        "S. Ramírez, FastAPI Documentation, https://fastapi.tiangolo.com.",
        "Solargis s.r.o., Solargis data and methodology documentation, https://solargis.com.",
        "Microsoft, TypeScript Handbook, https://www.typescriptlang.org/docs.",
        "R. C. Martin, Clean Architecture: A Craftsman's Guide to Software Structure and Design, Prentice Hall, 2017.",
        "D. Thomas and A. Hunt, The Pragmatic Programmer, 20th Anniversary Edition, Addison-Wesley, 2019.",
        "United Nations, Sustainable Development Goals, https://sdgs.un.org/goals.",
    ]
    for i, r in enumerate(refs, 1):
        p = L.para(d, f"[{i}] {r}", align=WD_ALIGN_PARAGRAPH.LEFT, space_after=4)
        p.paragraph_format.left_indent = Cm(0.9)
        p.paragraph_format.first_line_indent = Cm(-0.9)


def build_body(report_title, pages=None):
    d = L.new_document()
    cover(d, report_title)
    L.new_section(d, fmt="lowerRoman", start=1)
    index_page(d, REPORT1_INDEX, pages, first=True)
    certificate_page(d)
    declaration(d)
    acknowledgement(d)
    about_company(d)
    front_lists(d)
    L.new_section(d, fmt="decimal", start=1)
    chapter1(d)
    _drop_break_after_section(d)
    chapter2(d); chapter3(d); chapter4(d); chapter5(d); chapter6(d)
    weekly_summary(d); learning_outcomes(d); references(d)
    return d


def _drop_break_after_section(d):
    for p in d.paragraphs:
        if p.text.startswith("CHAPTER 1"):
            p.paragraph_format.page_break_before = False
            break

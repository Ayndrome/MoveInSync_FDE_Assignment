#!/usr/bin/env python3

import sys, re
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image, HRFlowable
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


COLORS = {
    "primary": "#000000",    # Near Black (Titles, H1, Table Headers)
    "secondary": "#475569",  # Slate (H2, Callout text)
    "chart": "#64748B",      # Muted Slate (Complementary chart color)
    "body": "#374151",       # Dark Grey (Body text for high readability)
    "muted": "#6B7280",      # Medium Grey (Footers, minor text)
    "grid": "#E5E7EB",       # Light Grey (Table and chart grids)
    "pale": "#F3F4F6",       # Very Light Grey (Callout backgrounds)
    "row_alt": "#F9FAFB",    # Faint Grey (Table alternate rows)
    "white": "#FFFFFF",
}

FONT_REGULAR = None
FONT_BOLD = None
FONT_ITALIC = None
REPORT_DATE = "12 October 2026"
CHART_FILENAME = "moveinsync_issue_distribution.png"

def esc(x):
    if pd.isna(x): return ""
    return str(x).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def build_report(xlsx, output_pdf):
    global BODY, BOLD, ITALIC
    BODY, BOLD, ITALIC = "Helvetica", "Helvetica-Bold", "Helvetica-Oblique"
    for name, path in [("MISRegular", FONT_REGULAR), ("MISBold", FONT_BOLD), ("MISItalic", FONT_ITALIC)]:
        if path:
            if not Path(path).exists(): raise FileNotFoundError(f"Font not found: {path}")
            pdfmetrics.registerFont(TTFont(name, path))
            if name == "MISRegular": BODY = name
            elif name == "MISBold": BOLD = name
            else: ITALIC = name
    C = {k: colors.HexColor(v) for k, v in COLORS.items()}
    S = getSampleStyleSheet()
    

    defs = [
        ("CoverTitle", "Title", BOLD, 26, 32, C["primary"], 12),
        ("SubtitleCustom", "Normal", BODY, 12, 18, C["muted"], 10),
        ("Section", "Heading1", BOLD, 16, 22, C["primary"], 12),
        ("Subsection", "Heading2", BOLD, 13, 18, C["secondary"], 6),
        ("BodyCustom", "BodyText", BODY, 10, 16, C["body"], 8),
        ("SmallCustom", "BodyText", BODY, 8.5, 13, C["muted"], 4),
        ("QuoteCustom", "BodyText", ITALIC, 9.5, 15, C["secondary"], 8),
        ("TableHeadCustom", "BodyText", BOLD, 9.5, 13, C["white"], 0),
        ("TableCellCustom", "BodyText", BODY, 9.5, 14, C["body"], 0),
        ("TableCellSmall", "BodyText", BODY, 8.5, 13, C["body"], 0),
    ]
    for name, parent, font, size, leading, color, after in defs:
        S.add(ParagraphStyle(name=name, parent=S[parent], fontName=font, fontSize=size,
            leading=leading, textColor=color, spaceAfter=after, alignment=TA_LEFT,
            keepWithNext=name in ("Section", "Subsection")))
            
    S.add(ParagraphStyle(name="Callout", parent=S["BodyText"], fontName=BODY, fontSize=10.5,
        leading=16, textColor=C["secondary"], backColor=C["pale"], 
        borderPadding=14, spaceBefore=10, spaceAfter=14, leftIndent=12, rightIndent=12))
        
    def P(t, style="BodyCustom"): return Paragraph(t, S[style])
    def cell(t, small=False): return Paragraph(esc(t), S["TableCellSmall" if small else "TableCellCustom"])
    
    def table(rows, widths):
        t = Table(rows, colWidths=widths, repeatRows=1, hAlign="LEFT")
        cmds = [("VALIGN",(0,0),(-1,-1),"TOP"),("LEFTPADDING",(0,0),(-1,-1),8),
                ("RIGHTPADDING",(0,0),(-1,-1),8),("TOPPADDING",(0,0),(-1,-1),6),
                ("BOTTOMPADDING",(0,0),(-1,-1),6),("GRID",(0,0),(-1,-1),0.5,C["grid"]),
                ("BACKGROUND",(0,0),(-1,0),C["primary"])]
        if len(rows)>1: cmds.append(("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,C["row_alt"]]))
        t.setStyle(TableStyle(cmds)); return t

    df = pd.read_excel(xlsx, sheet_name="Cleaned_Reviews")
    required = {"review_id","primary_issue","duplicate_status","star_rating","platform","review_text_original"}
    if not required.issubset(df.columns): raise ValueError(f"Missing columns: {sorted(required-set(df.columns))}")
    data = df[df["duplicate_status"] != "confirmed_duplicate"].copy()
    counts = data["primary_issue"].fillna("unclassifiable_pending_text").value_counts()
    n, total = len(data), len(df)
    confirmed = int((df["duplicate_status"] == "confirmed_duplicate").sum())
    suspected = int((df["duplicate_status"] == "suspected_duplicate").sum())
    platforms = df["platform"].value_counts().to_dict()
    chart_path = Path(output_pdf).parent / CHART_FILENAME
    chart_data = counts.head(10).sort_values()
    
    fig, ax = plt.subplots(figsize=(6.7, 4.0)) 
    bars = ax.barh(chart_data.index.str.replace("_"," "), chart_data.values, color=COLORS["chart"])
    ax.set_xlabel("Reviews (confirmed duplicate excluded)", fontsize=9, color=COLORS["muted"], labelpad=10) 
    ax.set_title("Most common primary complaint categories", fontsize=11, fontweight='bold', color=COLORS["primary"], pad=15)
    ax.grid(axis="x", alpha=0.5, color=COLORS["grid"])
    ax.set_axisbelow(True)
    ax.tick_params(colors=COLORS["body"], labelsize=8)
    
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color(COLORS["grid"])
    ax.spines['bottom'].set_color(COLORS["grid"])
    
    for bar, val in zip(bars, chart_data.values): 
        ax.text(val+.7,bar.get_y()+bar.get_height()/2,str(int(val)),va="center",fontsize=8, color=COLORS["body"])
    ax.set_xlim(0,max(chart_data.values)*1.18 if len(chart_data) else 1); fig.tight_layout()
    fig.savefig(chart_path,dpi=180,bbox_inches="tight"); plt.close(fig)

    def quote(rid, max_chars=220):
        rows=data[data.review_id.astype(str)==rid]
        if rows.empty: return None
        r=rows.iloc[0]; title=esc(r.get("title","")); body=re.sub(r"\s+"," ",str(r.get("review_text_original",""))).strip()
        if body.lower() in ("nan","none"): body=""
        if len(body)>max_chars: body=body[:max_chars].rsplit(" ",1)[0]+"…"
        return P(f"<b>{esc(rid)} - {title}</b><br/>{esc(body)}","QuoteCustom")

    story=[Spacer(1,10*mm),P("MoveInSync","CoverTitle"),
      P("Why are users rating the commute app poorly and what should be fixed first?","SubtitleCustom"),
      Spacer(1,3*mm),HRFlowable(width="100%",thickness=1.5,color=C["secondary"],spaceBefore=5,spaceAfter=15),
      P(f"<b>Recommendation brief</b> | Prepared {REPORT_DATE} | Evidence base: {total} collected one-star reviews"),Spacer(1,5*mm),
      P("The clearest signal in this sample is a cluster of complaints around the reliability of the commute journey: knowing where the vehicle is and when it will arrive, trusting that a booking will persist and be dispatched correctly, and being able to open and access the app when needed. These failures matter because the app mediates a time-sensitive employee transport service.","Callout"),
      P("Executive summary","Section"),
      P(f"After excluding {confirmed} confirmed duplicate, the working analysis contains <b>{n} review records</b>. Location/ETA/routing leads ({int(counts.get('location_eta_routing',0))}; {counts.get('location_eta_routing',0)/n:.1%}), followed by app reliability ({int(counts.get('app_reliability',0))}; {counts.get('app_reliability',0)/n:.1%}), booking/scheduling ({int(counts.get('booking_scheduling',0))}; {counts.get('booking_scheduling',0)/n:.1%}), login/access ({int(counts.get('login_access',0))}; {counts.get('login_access',0)/n:.1%}), and feature usability ({int(counts.get('feature_usability',0))}; {counts.get('feature_usability',0)/n:.1%}). Together, these five categories account for {int(counts.head(5).sum())} of {n} records ({counts.head(5).sum()/n:.1%})."),
      P("<b>Decision:</b> start with a cross-functional reliability sprint focused on trip-state correctness and live trip information, while running app-crash and login diagnostics in parallel."),
      P("What to do first","Subsection")]
      
    pr=[[P("Priority","TableHeadCustom"),P("Focus","TableHeadCustom"),P("Why it comes first","TableHeadCustom")],
        [cell("1 - Immediate"),cell("Booking + dispatch state integrity"),cell("A booking that disappears, changes, or is marked no-show incorrectly can disrupt an employee's commute. Instrument and reconcile each state transition.")],
        [cell("2 - Immediate"),cell("Live location, ETA and driver contact"),cell("Largest complaint category. Stale location and uncertain pickup timing erode trust. Show freshness and align driver identity/contact data.")],
        [cell("3 - Parallel"),cell("Crash-free startup + login/OTP"),cell("Crashes and OTP failures block the whole journey. Segment by app version, OS/device and authentication step before selecting a fix.")],
        [cell("4 - Next"),cell("Operational recovery and support"),cell("Give users a clear way to report trip issues, reach a human, and dispute cancellation/no-show outcomes.")],
        [cell("5 - Follow-through"),cell("Usability, permissions and feedback"),cell("Reduce friction after core journey reliability is under control.")]]
    story += [table(pr,[24*mm,46*mm,100*mm]),Spacer(1,4*mm),P("These are frequencies within a collected one-star-review sample, not estimates of the percentage of all users affected.","SmallCustom"),
      P("1. Dataset and method","Section"),P(f"The workbook contains {total} records: {platforms.get('Google Play Store',0)} Google Play reviews and {platforms.get('Apple App Store',0)} Apple App Store reviews. The collector confirmed all were one-star reviews. One confirmed duplicate was excluded, leaving {n} records {suspected} suspected duplicates remain included."),
      P("Primary issue labels were counted, and representative review text was checked against leading categories. Recommendations are proposed actions and hypotheses the spreadsheet alone cannot establish technical root causes."),
      P("What this sample can and cannot tell us","Subsection")]
    method=[[P("Can support","TableHeadCustom"),P("Cannot establish by itself","TableHeadCustom")],
      [cell("Recurring complaint themes among collected one-star reviews."),cell("Overall app-store rating distribution or proportion of all users affected.")],
      [cell("Examples of failed tracking, booking, login and support journeys."),cell("Whether each complaint is current, isolated, device-specific or already resolved.")],
      [cell("A prioritized validation plan and candidate fixes."),cell("Root cause or causal impact on star ratings.")]]
    story += [table(method,[85*mm,85*mm]),Spacer(1,3*mm),P("Data-quality notes","Subsection"),
      P("Individual review URLs and collection dates are unavailable some normalized dates may be inferred. Do not claim a reliable time trend. Platform comparisons are descriptive only. Two suspected duplicates remain included, so counts may shift if they are confirmed."),
      P("2. What the reviews say","Section"),P("The top five categories account for nearly three quarters of the working sample. The strongest near-term opportunity is to make the core commute journey dependable before prioritizing lower-frequency polish requests."),Image(str(chart_path),width=170*mm,height=101.5*mm)]
    readings={"location_eta_routing":"Vehicle visibility, stale/incorrect location, ETA and route/pickup sequencing.","app_reliability":"Crashes, freezes, launch failures and update/device compatibility reports.","booking_scheduling":"Missing or changed trips, booking edits not sticking, duplicate or incorrect trip states.","login_access":"OTP delivery, phone entry and account-access failures.","feature_usability":"Confusing flows, search/date selection and workflow friction.","operations_availability":"Late pickups, vehicle availability and dispatch/service execution.","driver_communication":"Incorrect driver details, failed contact or blocked communication.","support_and_policy":"Unresponsive support, cancellation restrictions and disputed no-shows.","feedback_ux":"Feedback prompts or interruptions.","address_mapping":"Incorrect or difficult-to-change pickup/drop addresses.","app_performance_size":"Slow performance or app-size concerns.","permissions_privacy":"Complaints about permissions and access controls.","safety_trust":"Safety-related concerns low frequency does not mean low severity.","parking_seat_booking":"Seat/parking booking friction.","other":"Miscellaneous reports."}
    issue_rows=[[P("Primary issue","TableHeadCustom"),P("Reviews","TableHeadCustom"),P(f"Share of {n}","TableHeadCustom"),P("Interpretation","TableHeadCustom")]]
    for cat,count in counts.items(): issue_rows.append([cell(cat.replace("_"," ")),cell(str(int(count))),cell(f"{count/n:.1%}"),cell(readings.get(cat,""),True)])
    story += [table(issue_rows,[40*mm,18*mm,22*mm,90*mm]),Spacer(1,3*mm),P("Counts use the primary classification. Secondary labels can overlap and should not be added as distinct reviews.","SmallCustom"),PageBreak(),P("3. Evidence from the review text","Section")]
    evidence=[("Location, ETA and routing",["MIS-043","MIS-017"],"Test location freshness, dispatch updates, ETA calculation and whether passenger, driver and transport desk see consistent trip state."),
      ("Booking and scheduling",["MIS-013","MIS-034"],"Trace the flow from user action through persisted booking, dispatch allocation and driver visibility."),
      ("Reliability and access",["MIS-010","MIS-004"],"Segment crash and authentication funnels by version, OS/device, network and failure step before selecting a fix."),
      ("Operational recovery and safety",["MIS-014","GP-085"],"Provide a recovery path for failed pickups. Escalate specific, credible safety reports regardless of frequency do not treat allegations as proven facts.")]
    for heading,ids,txt in evidence:
        story.append(P(heading,"Subsection"))
        for rid in ids:
            q=quote(rid)
            if q: story.append(q)
        story.append(P(txt))
    story += [PageBreak(),P("4. Prioritized action plan","Section"),P("First 72 hours: establish facts and contain obvious failures","Subsection")]
    for item in ["<b>Assign one incident owner</b> across app engineering, dispatch/operations and support review recent incidents daily.","<b>Instrument the journey</b>: OTP requested → verified → booking created/edited/cancelled → dispatch assigned → driver data refreshed → trip completed/no-show.","<b>Audit state mismatches</b> between app-visible trip state and backend booking, dispatch and driver records.","<b>Review crash and OTP telemetry</b> by app version, OS/device and failure step. Roll back only if a regression is confirmed.","<b>Triage safety complaints</b> and disputed no-show/cancellation cases."]:
        story.append(P("• "+item))
    story.append(P("Days 4–14: fix highest-impact failure modes","Subsection"))
    ws=[[P("Workstream","TableHeadCustom"),P("Action","TableHeadCustom"),P("Validation metric","TableHeadCustom")],
      [cell("Trip-state integrity"),cell("Make booking transitions idempotent reconcile booking and dispatch alert on disappearing/duplicate trips.",True),cell("Booking mutation success state mismatch disputed no-show rate.",True)],
      [cell("Tracking and ETA"),cell("Track location freshness show last-updated time when stale compare ETA with actual arrival.",True),cell("Location freshness ETA error wrong-driver reports on-time pickup.",True)],
      [cell("Crash and login"),cell("Fix leading crash signatures measure OTP funnel provide clear retry states.",True),cell("Crash-free sessions startup success OTP conversion.",True)],
      [cell("Support recovery"),cell("Provide in-trip help, case ID and status route failed pickups and disputes to a responsible queue.",True),cell("Time to response and resolution repeat contacts.",True)]]
    story += [table(ws,[30*mm,70*mm,70*mm]),Spacer(1,3*mm),P("Days 15–30: remove recurring friction","Subsection")]
    for item in ["Fix usability blockers after core reliability is measured test task completion, not visual preference alone.","Review permissions against actual feature needs and explain requests clearly.","Audit cancellation/no-show rules, cutoffs, exceptions and dispute paths with operations.","Monitor new one-star review volume and issue themes by release after changes ship."]:
        story.append(P("• "+item))
    story += [PageBreak(),P("5. Metrics and decision gates","Section"),P("Set a baseline in the first 72 hours before promising numerical improvement. Measure the employee commute outcome, not store ratings alone." )]
    met=[[P("Metric","TableHeadCustom"),P("Definition","TableHeadCustom"),P("Decision use","TableHeadCustom")],
      [cell("Crash-free sessions"),cell("Sessions without crashes, segmented by version and OS/device.",True),cell("Find regressions and verify fixes.",True)],
      [cell("OTP funnel"),cell("OTP requested, sent, delivered where measurable, entered and verified.",True),cell("Separate delivery from UI/input/verification failures.",True)],
      [cell("Booking-state consistency"),cell("Share of booking actions whose final state matches dispatch and driver-facing state.",True),cell("Confirm trips persist and edits/cancellations apply correctly.",True)],
      [cell("Location freshness / ETA error"),cell("Age of last location update and difference between predicted and actual pickup.",True),cell("Detect stale tracking and misleading ETAs.",True)],
      [cell("Pickup reliability"),cell("On-time pickup, assignment delay, wrong-driver contact and disputed no-show rate.",True),cell("Connect software state to operations.",True)],
      [cell("Support recovery"),cell("Time to response/resolution, repeat contacts and reopened cases.",True),cell("Verify users can recover when a ride fails.",True)],
      [cell("Review outcomes"),cell("New review volume and issue labels by release/cohort alongside rating distribution.",True),cell("Check pain trends without claiming causality.",True)]]
    story += [table(met,[35*mm,70*mm,65*mm]),Spacer(1,4*mm),P("Proposed release gate","Subsection"),P("Do not declare success just because a fix ships. Require targeted telemetry to improve, no material regression in adjacent journey steps, and fewer matching complaints in the next review cohort. Where telemetry is unavailable, instrument first and use a time-boxed manual audit."),P("Risks and trade-offs","Subsection")]
    risks=[[P("Risk","TableHeadCustom"),P("Mitigation","TableHeadCustom")],
      [cell("Largest category focus may miss blocking crashes/login."),cell("Prioritize by frequency and journey severity run diagnostics in parallel.",True)],
      [cell("Treating anecdotes as proof of root cause."),cell("Confirm symptoms with telemetry, logs and operations records.",True)],
      [cell("Focusing on UI while operations remain inconsistent."),cell("Joint ownership across mobile, backend/dispatch and operations.",True)],
      [cell("Low-count safety issues being deprioritized."),cell("Escalate specific safety reports independently of frequency.",True)]]
    story += [table(risks,[75*mm,95*mm]),PageBreak(),P("6. Conclusion and methodology note","Section"),
      P("Make the core commute journey trustworthy: booking state must persist, trip assignment and driver details must agree, location and ETA must be fresh enough to act on, and the app must remain accessible through reliable startup and login. Together, these determine whether an employee can plan, find and board the correct ride with confidence."),
      P("Start with trip-state reconciliation and location/ETA instrumentation, run crash and OTP diagnostics in parallel, and give users a reliable recovery path. Establish baselines in week one, fix the best-supported failure modes next, and validate outcomes over the rest of the month.","Callout"),
      P("Methodology and tools used","Subsection"),P("The supplied cleaned_reviews.xlsx workbook was analyzed with Python/pandas. Primary-issue frequencies and percentages were calculated after excluding the confirmed duplicate. Representative review text was checked against category labels. Matplotlib produced the chart and ReportLab generated the PDF. No external sources were used to substantiate product defects recommendations are hypotheses to validate against internal telemetry and operations data."),
      P("Limitations","Subsection"),P("This is a convenience sample of one-star reviews, not a random sample of all users. Ratings were confirmed by the collector. Individual review URLs and collection dates can be avail in Google Play Store and Apple App Store respectively. some normalized dates may be inferred. Two suspected duplicates remain included. Reviews may be subjective, incomplete or refer to historical versions. This report ranks observed complaint themes it does not claim verified root causes or estimate overall customer prevalence."),Spacer(1,7*mm),HRFlowable(width="100%",thickness=.7,color=C["grid"],spaceBefore=3,spaceAfter=7),P(f"Prepared for the MoveInSync Forward Deployed Engineer candidate assignment · {REPORT_DATE}","SmallCustom")]

    def footer(canvas, doc):
        canvas.saveState(); w,h=A4; canvas.setStrokeColor(C["grid"]); canvas.setLineWidth(.5)
        canvas.line(20*mm,15*mm,w-20*mm,15*mm); canvas.setFont(BODY,8); canvas.setFillColor(C["muted"])
        canvas.drawString(20*mm,10*mm,"MoveInSync review analysis")
        canvas.drawRightString(w-20*mm,10*mm,str(doc.page)); canvas.restoreState()
        
    doc=SimpleDocTemplate(str(output_pdf),pagesize=A4,rightMargin=20*mm,leftMargin=20*mm,topMargin=20*mm,bottomMargin=20*mm,title="MoveInSync Review Analysis and Recommendations",author="Candidate analysis")
    doc.build(story,onFirstPage=footer,onLaterPages=footer)
    print(f"Created {output_pdf}\nAnalyzed {n}/{total} records excluded {confirmed} confirmed duplicate(s); retained {suspected} suspected duplicate(s).")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python generate_moveinsync_report.py cleaned_reviews.xlsx [output.pdf]")
        raise SystemExit(2)
    input_file=Path(sys.argv[1]).expanduser()
    output_file=Path(sys.argv[2]).expanduser() if len(sys.argv)>2 else Path("MoveInSync_Review_Analysis_and_Recommendations.pdf")
    build_report(input_file,output_file)


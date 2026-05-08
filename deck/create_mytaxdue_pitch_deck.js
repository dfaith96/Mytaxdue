const fs = require("fs");
const path = require("path");

const pptxgen = require(path.join(process.env.APPDATA, "npm", "node_modules", "pptxgenjs"));

const root = path.resolve(__dirname, "..");
const primaryOutPath = path.join(__dirname, "mytaxdue_Pitch_Deck.pptx");
const fallbackOutPath = path.join(__dirname, "mytaxdue_Pitch_Deck_Updated.pptx");

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "mytaxdue";
pptx.company = "mytaxdue";
pptx.subject = "mytaxdue AI tax compliance assistant pitch deck";
pptx.title = "mytaxdue Pitch Deck";
pptx.lang = "en-US";
pptx.theme = {
  headFontFace: "Aptos Display",
  bodyFontFace: "Aptos",
  lang: "en-US"
};
pptx.margin = 0;

const C = {
  ink: "18202D",
  muted: "5C6673",
  paper: "F6F5EE",
  panel: "FFFFFF",
  line: "DCE2E0",
  forest: "0D5548",
  green: "15803D",
  mint: "DFF2E7",
  amber: "C98C16",
  clay: "BF5B3F",
  sky: "E5F1F4"
};

function svgData(name) {
  const svg = fs.readFileSync(path.join(root, "assets", name), "utf8");
  return `data:image/svg+xml;base64,${Buffer.from(svg).toString("base64")}`;
}

function addBackground(slide, dark = false) {
  slide.background = { color: dark ? C.ink : C.paper };
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 13.333,
    h: 0.12,
    fill: { color: dark ? C.amber : C.forest },
    line: { color: dark ? C.amber : C.forest }
  });
}

function addLogo(slide, dark = false) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.55,
    y: 0.38,
    w: 0.52,
    h: 0.52,
    rectRadius: 0.06,
    fill: { color: dark ? C.panel : C.forest },
    line: { color: dark ? C.panel : C.forest }
  });
  slide.addText("MTD", {
    x: 0.63,
    y: 0.55,
    w: 0.36,
    h: 0.13,
    fontSize: 7.5,
    bold: true,
    color: dark ? C.forest : C.panel,
    align: "center",
    margin: 0
  });
  slide.addText("mytaxdue", {
    x: 1.15,
    y: 0.5,
    w: 1.5,
    h: 0.22,
    fontFace: "Aptos Display",
    fontSize: 16,
    bold: true,
    color: dark ? C.panel : C.ink,
    margin: 0
  });
}

function addFooter(slide, number, dark = false) {
  slide.addText(`mytaxdue / ${String(number).padStart(2, "0")}`, {
    x: 0.62,
    y: 7.1,
    w: 3.6,
    h: 0.18,
    fontSize: 8,
    color: dark ? "DCE4E4" : C.muted,
    margin: 0
  });
}

function addTitle(slide, eyebrow, title, body, dark = false) {
  slide.addText(eyebrow.toUpperCase(), {
    x: 0.72,
    y: 1.05,
    w: 4.5,
    h: 0.22,
    fontSize: 9,
    bold: true,
    color: dark ? C.amber : C.clay,
    charSpace: 1.1,
    margin: 0
  });
  slide.addText(title, {
    x: 0.68,
    y: 1.36,
    w: 5.9,
    h: 1.04,
    fontFace: "Aptos Display",
    fontSize: 34,
    bold: true,
    color: dark ? C.panel : C.ink,
    fit: "shrink",
    margin: 0
  });
  if (body) {
    slide.addText(body, {
      x: 0.72,
      y: 2.52,
      w: 5.25,
      h: 0.82,
      fontSize: 14,
      color: dark ? "DCE4E4" : C.muted,
      fit: "shrink",
      margin: 0
    });
  }
}

function card(slide, x, y, w, h, title, body, accent = C.forest) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.05,
    fill: { color: C.panel },
    line: { color: C.line, width: 1 }
  });
  slide.addShape(pptx.ShapeType.rect, {
    x,
    y,
    w,
    h: 0.08,
    fill: { color: accent },
    line: { color: accent }
  });
  slide.addText(title, {
    x: x + 0.2,
    y: y + 0.22,
    w: w - 0.4,
    h: 0.26,
    fontSize: 15,
    bold: true,
    color: C.ink,
    margin: 0
  });
  slide.addText(body, {
    x: x + 0.2,
    y: y + 0.62,
    w: w - 0.4,
    h: h - 0.78,
    fontSize: 10.5,
    color: C.muted,
    fit: "shrink",
    margin: 0
  });
}

function addBullets(slide, x, y, items, dark = false) {
  items.forEach((item, index) => {
    const top = y + index * 0.72;
    slide.addShape(pptx.ShapeType.ellipse, {
      x,
      y: top + 0.08,
      w: 0.16,
      h: 0.16,
      fill: { color: dark ? C.amber : C.green },
      line: { color: dark ? C.amber : C.green }
    });
    slide.addText(item, {
      x: x + 0.28,
      y: top,
      w: 4.95,
      h: 0.42,
      fontSize: 13,
      color: dark ? "E9F0EE" : C.ink,
      fit: "shrink",
      margin: 0
    });
  });
}

function pill(slide, x, y, text, color, width = 1.3) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w: width,
    h: 0.38,
    rectRadius: 0.12,
    fill: { color },
    line: { color }
  });
  slide.addText(text, {
    x: x + 0.08,
    y: y + 0.11,
    w: width - 0.16,
    h: 0.14,
    fontSize: 8.5,
    bold: true,
    color: color === C.mint ? C.forest : C.panel,
    align: "center",
    margin: 0
  });
}

function slide1() {
  const slide = pptx.addSlide();
  addBackground(slide, true);
  addLogo(slide, true);
  addTitle(
    slide,
    "Pitch deck",
    "mytaxdue: AI tax planning and compliance assistant.",
    "We help Nigerian taxpayers, SMEs, and accountants estimate tax exposure, organise records, find possible reliefs, and prepare cleaner filing packs.",
    true
  );
  slide.addImage({ data: svgData("tax-dashboard.svg"), x: 6.25, y: 1.04, w: 6.45, h: 4.15 });
  pill(slide, 0.72, 3.66, "Tax estimate", C.green, 1.45);
  pill(slide, 2.34, 3.66, "Document vault", C.amber, 1.55);
  pill(slide, 4.1, 3.66, "Expert review", C.clay, 1.55);
  slide.addText("Planning estimates only. Final filings require professional or authority review.", {
    x: 0.72,
    y: 5.92,
    w: 7.0,
    h: 0.42,
    fontSize: 15,
    color: "DCE4E4",
    margin: 0
  });
  addFooter(slide, 1, true);
}

function slide2() {
  const slide = pptx.addSlide();
  addBackground(slide);
  addLogo(slide);
  addTitle(slide, "Problem", "Tax compliance is still confusing and document-heavy.", "Individuals, freelancers, SMEs, and accountants lose time collecting records, estimating obligations, and preparing compliant submissions.");
  card(slide, 6.55, 1.05, 2.0, 1.55, "Scattered records", "Receipts, invoices, PAYE records, WHT notes, and relief evidence sit in different places.", C.clay);
  card(slide, 8.8, 1.05, 2.0, 1.55, "Late visibility", "Taxpayers often discover what they owe too close to the deadline.", C.amber);
  card(slide, 11.05, 1.05, 1.85, 1.55, "Manual review", "Accountants spend too much time cleaning inputs before advising clients.", C.green);
  slide.addShape(pptx.ShapeType.line, { x: 1.0, y: 4.72, w: 10.9, h: 0, line: { color: C.line, width: 2 } });
  ["Earn", "Collect records", "Estimate", "Review", "File"].forEach((label, i) => {
    const x = 1.0 + i * 2.72;
    slide.addShape(pptx.ShapeType.ellipse, {
      x,
      y: 4.4,
      w: 0.64,
      h: 0.64,
      fill: { color: i === 1 ? C.clay : C.panel },
      line: { color: i === 1 ? C.clay : C.green, width: 2 }
    });
    slide.addText(label, { x: x - 0.3, y: 5.2, w: 1.35, h: 0.25, fontSize: 10.5, bold: true, color: C.ink, align: "center", margin: 0 });
  });
  slide.addText("The bottleneck is turning messy records into a trusted tax position.", {
    x: 3.1,
    y: 5.82,
    w: 7.0,
    h: 0.34,
    fontSize: 18,
    bold: true,
    color: C.forest,
    align: "center",
    margin: 0
  });
  addFooter(slide, 2);
}

function slide3() {
  const slide = pptx.addSlide();
  addBackground(slide);
  addLogo(slide);
  addTitle(slide, "Solution", "A tax planning workspace before the deadline.", "mytaxdue organises financial records, runs planning estimates, suggests missing documents, and creates expert-ready filing packs.");
  card(slide, 6.6, 1.15, 2.85, 1.55, "1. Capture", "Add income, expenses, tax credits, relief evidence, and documents.", C.green);
  card(slide, 9.85, 1.15, 2.85, 1.55, "2. Estimate", "Generate a transparent planning reserve with assumptions and gaps.", C.amber);
  card(slide, 6.6, 3.15, 2.85, 1.55, "3. Optimise", "Find missing documents and possible relief categories for review.", C.clay);
  card(slide, 9.85, 3.15, 2.85, 1.55, "4. Review", "Send a clean tax pack to an accountant before filing.", C.forest);
  slide.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 4.78, w: 5.0, h: 1.12, rectRadius: 0.08, fill: { color: C.ink }, line: { color: C.ink } });
  slide.addText("Outcome: better preparation, fewer missing records, and earlier visibility into possible tax due.", {
    x: 1.08,
    y: 5.1,
    w: 4.45,
    h: 0.44,
    fontSize: 14,
    color: C.panel,
    fit: "shrink",
    margin: 0
  });
  addFooter(slide, 3);
}

function slide4() {
  const slide = pptx.addSlide();
  addBackground(slide, true);
  addLogo(slide, true);
  addTitle(slide, "Product", "Tax planning dashboard for taxpayers and accountants.", "A single workspace for estimates, documents, compliance tasks, and review packs.", true);
  slide.addImage({ data: svgData("tax-dashboard.svg"), x: 6.05, y: 0.94, w: 6.65, h: 4.28 });
  addBullets(slide, 0.92, 3.72, [
    "Estimate tax due from income, expenses, reliefs, and credits.",
    "Organise records into a document vault.",
    "Track filing tasks and deadline readiness.",
    "Export a review pack for tax professionals."
  ], true);
  addFooter(slide, 4, true);
}

function slide5() {
  const slide = pptx.addSlide();
  addBackground(slide);
  addLogo(slide);
  addTitle(slide, "Workflow", "From scattered tax records to filing readiness.", "The platform turns everyday records into a clear tax preparation process.");
  slide.addImage({ data: svgData("tax-workflow.svg"), x: 6.55, y: 1.18, w: 5.5, h: 3.8 });
  addBullets(slide, 0.92, 3.55, [
    "Capture income and expense evidence.",
    "AI organises documents and highlights gaps.",
    "Planning estimate shows possible exposure.",
    "Accountant verifies the final position before filing."
  ]);
  addFooter(slide, 5);
}

function slide6() {
  const slide = pptx.addSlide();
  addBackground(slide);
  addLogo(slide);
  addTitle(slide, "Customers", "Built for people who need tax clarity before they file.", "The first customers are users with recurring tax obligations and accountants managing many clients.");
  card(slide, 6.6, 1.2, 2.9, 1.6, "Individuals", "PAYE workers and professionals who want visibility into tax exposure.", C.green);
  card(slide, 9.8, 1.2, 2.9, 1.6, "Freelancers", "Independent earners tracking invoices, WHT credits, and expenses.", C.amber);
  card(slide, 6.6, 3.2, 2.9, 1.6, "SMEs", "Small businesses preparing records for tax and accountant review.", C.clay);
  card(slide, 9.8, 3.2, 2.9, 1.6, "Accountants", "Professionals who need cleaner client intake and review workflows.", C.forest);
  addBullets(slide, 0.95, 3.65, [
    "Beachhead: freelancers, SME owners, and accountants in major Nigerian cities.",
    "Entry point: free planning estimate and document checklist.",
    "Expansion: review packs, client workspace, and compliance calendar."
  ]);
  addFooter(slide, 6);
}

function slide7() {
  const slide = pptx.addSlide();
  addBackground(slide, true);
  addLogo(slide, true);
  addTitle(slide, "Business model", "Freemium tax planning with paid automation and professional workflows.", "mytaxdue can monetise through personal plans, accountant workspaces, and expert review services.", true);
  card(slide, 0.85, 3.05, 2.85, 1.7, "Free planner", "Basic tax estimate, checklist, and deadline reminders.", C.green);
  card(slide, 3.98, 3.05, 2.85, 1.7, "Personal Plus", "Document vault, relief finder, estimate history, and export pack.", C.amber);
  card(slide, 7.1, 3.05, 2.85, 1.7, "Expert review", "Paid review with a tax professional before final filing.", C.clay);
  card(slide, 10.22, 3.05, 2.35, 1.7, "Accountant", "Client workspace, team access, and bulk preparation tools.", C.forest);
  slide.addText("Revenue grows with users, document volume, accountant seats, review requests, and future integrations.", {
    x: 1.2,
    y: 5.52,
    w: 10.9,
    h: 0.36,
    fontSize: 16,
    bold: true,
    color: C.panel,
    align: "center",
    margin: 0
  });
  addFooter(slide, 7, true);
}

function slide8() {
  const slide = pptx.addSlide();
  addBackground(slide);
  addLogo(slide);
  addTitle(slide, "Go-to-market", "Acquire users through tax season urgency and accountant partnerships.", "The product starts with a simple tax estimate and converts users into paid preparation workflows.");
  addBullets(slide, 0.92, 3.05, [
    "Launch free calculator and checklist for common taxpayer profiles.",
    "Partner with accountants and SME communities for trusted distribution.",
    "Use content around deadlines, documents, WHT, PAYE, and expense readiness.",
    "Convert active users to document vaults, review packs, and accountant plans."
  ]);
  slide.addShape(pptx.ShapeType.roundRect, { x: 7.2, y: 1.25, w: 4.8, h: 4.9, rectRadius: 0.06, fill: { color: C.panel }, line: { color: C.line } });
  ["Estimate", "Checklist", "Review pack", "Paid workflow"].forEach((label, i) => {
    const y = 1.78 + i * 0.93;
    const color = [C.green, C.amber, C.clay, C.forest][i];
    slide.addShape(pptx.ShapeType.ellipse, { x: 7.68, y, w: 0.42, h: 0.42, fill: { color }, line: { color } });
    slide.addText(label, { x: 8.35, y: y + 0.06, w: 2.5, h: 0.22, fontSize: 15, bold: true, color: C.ink, margin: 0 });
    if (i < 3) slide.addShape(pptx.ShapeType.line, { x: 7.89, y: y + 0.48, w: 0, h: 0.45, line: { color: C.line, width: 2 } });
  });
  addFooter(slide, 8);
}

function slide9() {
  const slide = pptx.addSlide();
  addBackground(slide);
  addLogo(slide);
  addTitle(slide, "Positioning", "More structured than a calculator, lighter than enterprise tax software.", "mytaxdue gives users a practical preparation layer that can connect to professionals before final filing.");
  const rows = [
    ["Capability", "Spreadsheet", "Basic calculator", "Manual accountant intake", "mytaxdue"],
    ["Tax estimate", "Manual", "Single-use", "Delayed", "Guided"],
    ["Document checklist", "Manual", "No", "Varies", "AI-assisted"],
    ["Client workspace", "No", "No", "Manual", "Built in"],
    ["Expert review pack", "Manual", "No", "Manual", "Exportable"]
  ];
  const x = 0.75;
  const y = 2.5;
  const colW = [2.35, 2.05, 2.1, 2.45, 1.95];
  rows.forEach((row, r) => {
    let left = x;
    row.forEach((cell, c) => {
      slide.addShape(pptx.ShapeType.rect, {
        x: left,
        y: y + r * 0.55,
        w: colW[c],
        h: 0.55,
        fill: { color: r === 0 ? C.ink : c === 4 ? "DFF2E7" : C.panel },
        line: { color: C.line, width: 0.6 }
      });
      slide.addText(cell, {
        x: left + 0.08,
        y: y + r * 0.55 + 0.16,
        w: colW[c] - 0.16,
        h: 0.16,
        fontSize: r === 0 ? 9 : 8.4,
        bold: r === 0 || c === 4,
        color: r === 0 ? C.panel : C.ink,
        fit: "shrink",
        margin: 0
      });
      left += colW[c];
    });
  });
  slide.addText("Differentiation: AI-assisted preparation plus human review readiness.", {
    x: 1.25,
    y: 6.0,
    w: 10.9,
    h: 0.32,
    fontSize: 15,
    bold: true,
    color: C.forest,
    align: "center",
    margin: 0
  });
  addFooter(slide, 9);
}

function slide10() {
  const slide = pptx.addSlide();
  addBackground(slide);
  addLogo(slide);
  addTitle(slide, "Trust", "Designed as a preparation tool with expert review.", "Tax decisions are sensitive. mytaxdue should make assumptions clear and encourage professional verification before filing.");
  addBullets(slide, 0.92, 3.2, [
    "Estimates show assumptions and explain what data was used.",
    "Documents are organised for auditability and review.",
    "Accountants can verify final positions before submission.",
    "User consent controls govern document sharing and exports."
  ]);
  slide.addShape(pptx.ShapeType.roundRect, { x: 7.05, y: 1.28, w: 4.95, h: 4.15, rectRadius: 0.08, fill: { color: C.panel }, line: { color: C.line } });
  slide.addText("Important boundary", { x: 7.55, y: 1.9, w: 4.0, h: 0.34, fontSize: 22, bold: true, color: C.ink, align: "center", margin: 0 });
  slide.addText("mytaxdue provides planning estimates and preparation workflows. Final filings should be reviewed by a qualified tax adviser or relevant authority.", {
    x: 7.65,
    y: 2.58,
    w: 3.8,
    h: 1.3,
    fontSize: 14,
    color: C.muted,
    align: "center",
    fit: "shrink",
    margin: 0
  });
  pill(slide, 8.25, 4.35, "Human review", C.forest, 2.3);
  addFooter(slide, 10);
}

function slide11() {
  const slide = pptx.addSlide();
  addBackground(slide, true);
  addLogo(slide, true);
  addTitle(slide, "Roadmap", "Build, validate, and integrate around trusted tax workflows.", "The next phase should prove demand, improve calculation coverage, and deepen accountant tooling.", true);
  card(slide, 0.9, 3.05, 2.65, 1.55, "MVP", "Calculator, document checklist, vault, contact intake, and export pack.", C.green);
  card(slide, 3.85, 3.05, 2.65, 1.55, "Accountants", "Client workspace, review queue, comments, and team access.", C.amber);
  card(slide, 6.8, 3.05, 2.65, 1.55, "Compliance", "Assumption logs, privacy controls, review workflows, and audit trails.", C.clay);
  card(slide, 9.75, 3.05, 2.65, 1.55, "Integrations", "Future links to payments, accounting tools, and official filing systems where available.", C.forest);
  slide.addText("Milestones: beta users, accountant partners, paid review conversion, and repeat tax-season usage.", {
    x: 1.05,
    y: 5.52,
    w: 11.2,
    h: 0.34,
    fontSize: 15,
    bold: true,
    color: C.panel,
    align: "center",
    margin: 0
  });
  addFooter(slide, 11, true);
}

function slide12() {
  const slide = pptx.addSlide();
  addBackground(slide);
  addLogo(slide);
  addTitle(slide, "The ask", "Help mytaxdue make tax preparation easier before the deadline.", "We are seeking beta users, accountant partners, and early backers to build a practical AI layer for Nigerian tax planning.");
  addBullets(slide, 0.94, 3.1, [
    "Pilot with freelancers, SME owners, and accountants.",
    "Validate estimate workflows and document checklist quality.",
    "Partner on professional review and client intake workflows.",
    "Support the build-out of trusted tax preparation automation."
  ]);
  slide.addShape(pptx.ShapeType.roundRect, { x: 7.05, y: 2.6, w: 5.0, h: 2.2, rectRadius: 0.08, fill: { color: C.ink }, line: { color: C.ink } });
  slide.addText("mytaxdue", { x: 7.55, y: 3.08, w: 4.0, h: 0.54, fontFace: "Aptos Display", fontSize: 32, bold: true, color: C.panel, align: "center", margin: 0 });
  slide.addText("Estimate. Organise. Review. Prepare.", { x: 7.55, y: 3.78, w: 4.0, h: 0.32, fontSize: 13, color: "DCE4E4", align: "center", margin: 0 });
  slide.addText("Contact form available on the mytaxdue website.", { x: 7.55, y: 4.22, w: 4.0, h: 0.34, fontSize: 10.5, color: "DCE4E4", align: "center", margin: 0 });
  addFooter(slide, 12);
}

[
  slide1,
  slide2,
  slide3,
  slide4,
  slide5,
  slide6,
  slide7,
  slide8,
  slide9,
  slide10,
  slide11,
  slide12
].forEach((build) => build());

pptx
  .writeFile({ fileName: primaryOutPath })
  .then(() => {
    console.log(primaryOutPath);
  })
  .catch((error) => {
    if (error && error.code === "EBUSY") {
      return pptx.writeFile({ fileName: fallbackOutPath }).then(() => {
        console.log(fallbackOutPath);
      });
    }
    throw error;
  });

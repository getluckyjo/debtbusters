const { Document, Packer, Paragraph, TextRun, AlignmentType, Footer, PageNumber,
        BorderStyle, TabStopType, TabStopPosition, PageBreak } = require('docx');
const fs = require('fs');
const FONT = "Times New Roman", SZ = 22;

const P = (t) => new TextRun({ text: t, bold: true, color: "C00000", font: FONT, size: SZ });
const T = (t) => new TextRun({ text: t, font: FONT, size: SZ });

function clause(num, runs) {
  return new Paragraph({ spacing: { after: 160, line: 300 }, indent: { left: 720, hanging: 720 },
    alignment: AlignmentType.JUSTIFIED, tabStops: [{ type: TabStopType.LEFT, position: 720 }],
    children: [ T(num + "\t"), ...(typeof runs === "string" ? [T(runs)] : runs) ] });
}
function lead(text) {
  return new Paragraph({ spacing: { after: 160, line: 300 }, indent: { left: 720 },
    alignment: AlignmentType.JUSTIFIED, children: [ T(text) ] });
}
function heading(num, text) {
  return new Paragraph({ spacing: { before: 280, after: 160 }, indent: { left: 720, hanging: 720 },
    tabStops: [{ type: TabStopType.LEFT, position: 720 }],
    children: [ new TextRun({ text: num + "\t" + text, bold: true, font: FONT, size: SZ, allCaps: true }) ] });
}
function centered(text, o = {}) {
  return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: o.after ?? 160 },
    children: [ new TextRun({ text, bold: o.bold ?? false, font: FONT, size: o.size ?? SZ }) ] });
}
function party(name, role) {
  return new Paragraph({ spacing: { after: 140 }, tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
    children: [ new TextRun({ text: name, bold: true, font: FONT, size: SZ }), T("\t"),
                new TextRun({ text: role, bold: true, font: FONT, size: SZ }) ] });
}
function sigLine(label) {
  return [ new Paragraph({ spacing: { before: 340, after: 60 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "000000", space: 1 } }, children: [T("")] }),
    new Paragraph({ spacing: { after: 200 }, children: [ new TextRun({ text: label, font: FONT, size: 20 }) ] }) ];
}
function fill(label) {
  return new Paragraph({ spacing: { after: 200 }, tabStops: [{ type: TabStopType.LEFT, position: 2600 }],
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "000000", space: 2 } },
    children: [ T(label), T("\t"), T("") ] });
}
function note(text) {
  return new Paragraph({ spacing: { after: 180, line: 300 }, alignment: AlignmentType.JUSTIFIED,
    children: [ new TextRun({ text, font: FONT, size: 20 }) ] });
}

const doc = new Document({
  creator: "Johannes Andries le Roux",
  title: "Supplementary Affidavit - Employment and Income",
  sections: [{
    properties: { page: { margin: { top: 1134, bottom: 1134, left: 1276, right: 1134 } } },
    footers: { default: new Footer({ children: [ new Paragraph({
      tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
      children: [ new TextRun({ text: "1st Resp. initial: ______   Commissioner of Oaths initial: ______", font: FONT, size: 18 }),
                  new TextRun({ text: "\t", font: FONT, size: 18 }),
                  new TextRun({ children: ["Page ", PageNumber.CURRENT, " of ", PageNumber.TOTAL_PAGES], font: FONT, size: 18 }) ] }) ] }) },
    children: [
      centered("IN THE MAGISTRATE'S COURT FOR THE DISTRICT OF WYNBERG", { bold: true, size: 22, after: 60 }),
      centered("HELD AT WYNBERG", { bold: true, size: 22, after: 200 }),
      new Paragraph({ spacing: { after: 240 }, children: [ T("Case Number: "), new TextRun({ text: "4652/2026", bold: true, font: FONT, size: SZ }) ] }),
      new Paragraph({ spacing: { after: 160 }, children: [ T("In the matter between:") ] }),
      party("SAUL ALLARDICE NCRDC4172", "APPLICANT"),
      new Paragraph({ spacing: { after: 100 }, children: [ T("and") ] }),
      party("JOHANNES ANDRIES LE ROUX", "1ST RESPONDENT"),
      new Paragraph({ spacing: { after: 160 }, children: [ new TextRun({ text: "ID NUMBER: 8901215147084", font: FONT, size: 20 }) ] }),
      party("FIRST NATIONAL BANK, A DIVISION OF FIRSTRAND BANK LIMITED", "2ND RESPONDENT"),
      party("WESBANK, A DIVISION OF FIRSTRAND BANK LIMITED", "3RD RESPONDENT"),
      new Paragraph({ spacing: { before: 200, after: 240 },
        border: { top: { style: BorderStyle.SINGLE, size: 6, color: "000000", space: 1 } }, children: [T("")] }),
      centered("SUPPLEMENTARY AFFIDAVIT", { bold: true, size: 26, after: 60 }),
      centered("Confirmation of employment and explanation of the fluctuation in monthly income", { size: 20, after: 260 }),

      new Paragraph({ spacing: { after: 200 }, children: [ T("I, the undersigned,") ] }),
      centered("JOHANNES ANDRIES LE ROUX", { bold: true, size: 26, after: 200 }),
      new Paragraph({ spacing: { after: 260 }, children: [ T("do hereby make oath and state that:") ] }),

      heading("1.", "Identity and standing of the deponent"),
      clause("1.1", [ T("I am an adult male with identity number 8901215147084. I am married out of community of property by antenuptial contract. I reside at "), P("40 Strubens Road, Mowbray, Cape Town, Western Cape"), T(" "), P("[CONFIRM STILL CURRENT]"), T(", and as such I reside within the area of jurisdiction of this Honourable Court.") ]),
      clause("1.2", "The facts set out in this affidavit are within my own personal knowledge, save where otherwise stated or where the contrary appears from the context, and are to the best of my knowledge and belief both true and correct."),
      clause("1.3", "I am the 1st Respondent and the consumer in this application for debt review brought on my behalf by the debt counsellor, Saul Allardice NCRDC4172, in terms of section 86 of the National Credit Act 34 of 2005. My reference with DebtBusters is wf_012447519."),
      clause("1.4", "I depose to this supplementary affidavit at the request of VHT Attorneys, the attorneys of record, following the request made at the previous hearing for updated proof of income, in order to confirm my employment and to explain the fluctuation in the monthly deposits reflected in the bank statements filed."),

      heading("2.", "My employment"),
      clause("2.1", [ T("I am employed. I am employed by "), P("[GET LUCKY GOLF — FULL REGISTERED NAME, e.g. Get Lucky Golf Club (Pty) Ltd, registration no. ...]"), T(" (“the Company”) in the position of "), P("[POSITION / JOB TITLE]"), T(", and have held that position since "), P("[MONTH AND YEAR]"), T(".") ]),
      clause("2.2", [ T("My agreed remuneration is R100,000.00 "), P("[gross / nett — DELETE ONE]"), T(" per month.") ]),
      clause("2.3", "I accordingly confirm that the description of my employment status as “EMPLOYED” in the confirmatory affidavit filed in support of this application is correct."),
      clause("2.4", [ T("I am also "), P("[a director and shareholder / a shareholder / the founder]"), T(" of the Company. "), P("[DELETE THIS PARAGRAPH IF YOU HOLD NO SHARES OR DIRECTORSHIP]") ]),
      clause("2.5", [ T("For the sake of completeness and accuracy, I record that in an email to the attorneys of record dated 8 September 2026 I described myself as “self-employed / business owner”. That description was imprecise. I used it because I founded the Company and "), P("[hold an interest in it]"), T(". The correct position, which I confirm under oath, is that I am employed by the Company and paid a salary by it, as set out above. I correct that description at the first opportunity and no inference adverse to the accuracy of my disclosures should be drawn from it.") ]),

      heading("3.", "Reasons for the fluctuation in my monthly income"),
      lead("Although my salary is a fixed monthly amount, the deposits reflected in my bank statements vary in both amount and date. That arises from the following:"),
      clause("3.1", "I am not always paid my full salary on the due date. The Company does not always have sufficient cash available on the salary date to pay my remuneration in full."),
      clause("3.2", [ T("When that occurs I am paid late, or I am paid a part only of my salary, or I am paid in more than one instalment during the course of the month. Any shortfall is carried forward and paid when the Company is able to do so. The effect is that a month in which a shortfall arises appears low, and a month in which arrears are settled appears correspondingly high. Over the past "), P("[12]"), T(" months this has occurred in approximately "), P("[NUMBER]"), T(" months.") ]),
      clause("3.3", [ T("The Company’s own revenue is project based and seasonal. It operates golf activations and hole-in-one challenges at golf courses, and earns its income from participant entry fees, brand sponsorships and partnership agreements. Sponsorship income is received irregularly and largely in arrears, and activity follows the golfing and corporate marketing calendar. That is the underlying cause of the timing difficulties described above.") ]),
      clause("3.4", [ T("In addition to my salary, I undertake consulting and advisory work in my own name, outside of my employment, "), P("[which my employer permits]"), T(". This work is engagement based: I invoice per engagement or per project, clients typically settle "), P("[30 to 60]"), T(" days after invoice and frequently later, and there is no fixed number of engagements in any month. The additional income is therefore irregular in both amount and timing, and some months yield none at all.") ]),
      clause("3.5", [ T("Over the "), P("[12]"), T(" months to "), P("[DATE]"), T(", the total amount I actually received was approximately R"), P("[AMOUNT]"), T(", being an average of approximately R"), P("[AMOUNT]"), T(" per month, with a lowest month of approximately R"), P("[AMOUNT]"), T(" and a highest month of approximately R"), P("[AMOUNT]"), T(".") ]),
      clause("3.6", "The fluctuation does not reflect concealed or additional income. The bank statements I have provided reflect the whole of my income from all sources."),

      heading("4.", "My ability to settle my debts"),
      clause("4.1", "I am unable to settle my debts in accordance with the original contractual terms of my credit agreements."),
      clause("4.2", [ T("My ability to pay is determined by what I actually receive and by my necessary living expenses, and not by my contractual salary. My necessary monthly expenses are approximately R"), P("[AMOUNT]"), T(", as more fully set out in the income and expenditure assessment conducted by the debt counsellor. "), P("[Consider annexing that assessment, or a current income and expenditure schedule, marked “JLR1”.]") ]),
      clause("4.3", "The original contractual instalments across the accounts under review substantially exceed what my income is able to sustain, and that is the reason I applied for debt review. I do not restate the precise total here, because I have raised a query with the debt counsellor regarding the accuracy of both the balances and the instalment figures reflected on the negotiation summary provided to me, and that query remains unresolved."),
      clause("4.4", "I am over-indebted as contemplated in section 79 of the National Credit Act 34 of 2005."),
      clause("4.5", [ T("I applied for debt review in January 2026 and have since "), P("[MONTH AND YEAR OF FIRST PAYMENT]"), T(" paid R17,000.00 per month to Hyphen PDA. "), P("[If correct: I have not missed a payment.]") ]),
      clause("4.6", [ T("Because no order of court has yet been granted, my credit providers continue to charge full contractual interest. The total reflected by the debt counsellor has increased from approximately R745,000.00 at the time of my application to approximately R1,204,976.00. I have raised a formal query regarding that increase. Approximately R487,000.00 of it arises on two accounts only, which I have reason to believe were captured at materially understated balances at the outset of the process, and which I have asked the debt counsellor to correct. I do not ask this Honourable Court to accept the increase as accurate, and I mention it only to explain the prejudice caused to me by the passage of time.") ]),
      clause("4.7", "I have paid R17,000.00 per month throughout and I am able and willing to continue doing so under a re-arrangement of my obligations. I am not seeking to avoid my obligations. I respectfully ask that my accounts be re-arranged in terms of section 87 of the National Credit Act 34 of 2005, on figures which have first been verified as correct."),

      heading("5.", "Conclusion"),
      clause("5.1", "I confirm that the contents of this affidavit are true and correct, and I make this statement in support of the debt review application referred to in paragraph 1.3 above, and in compliance with the request made at the previous hearing."),

      new Paragraph({ children: [T("")], spacing: { after: 120 } }),
      ...sigLine("DEPONENT — JOHANNES ANDRIES LE ROUX (ID 8901215147084)"),
      new Paragraph({ spacing: { after: 240 }, alignment: AlignmentType.JUSTIFIED,
        children: [ T("SIGNED and SWORN to before me at "), P("[PLACE]"), T(" on this "), P("[   ]"), T(" day of "), P("[MONTH]"), T(" 2026, the Deponent having acknowledged that he knows and understands the contents of this affidavit, which is deposed to in accordance with the regulations governing the administration of an oath as more fully set out in Government Notice R1258 of 21 July 1972, as amended by Government Notice 1648 dated 19 August 1977 and Government Notice 903 dated 10 July 1998.") ] }),
      ...sigLine("COMMISSIONER OF OATHS"),
      fill("Full names:"),
      fill("Rank / designation:"),
      fill("Ex officio (e.g. South African Police Service):"),
      fill("Station / office and address:"),
      fill("Official stamp and date:"),

      new Paragraph({ children: [ new PageBreak() ] }),
      centered("BEFORE YOU SIGN — NOTES, NOT PART OF THE AFFIDAVIT", { bold: true, size: 24, after: 240 }),
      note("Delete this entire page before you print the affidavit for signature."),
      note("1. You are EMPLOYED, not self-employed. The confirmatory affidavit already filed was correct. What needed correcting was your own email to Lyndon of 8 September describing yourself as “self-employed / business owner”, and paragraph 2.5 does that. Do not delete 2.5 — correcting it yourself is far better than having FNB’s counsel raise it on 26 October."),
      note("2. Expect the R100,000 salary to be questioned. A court will ask how someone on R100,000 a month can only pay R17,000 towards R1,240,177 of debt. Paragraphs 3.5 and 4.2 are what answer that: what you ACTUALLY received, and what your necessary expenses actually are. Complete them carefully and consider annexing the income and expenditure assessment. Leaving them vague is the single biggest risk in this document."),
      note("3. Paragraph 3.2 asks how many months in the past year involved a short or late payment. Get this from your bank statements rather than estimating — the statements are already before the court and the Magistrate can check."),
      note("4. Paragraph 2.4 — confirm whether you are a director, a shareholder, or both. If you hold neither, delete it. Paragraph 2.5 depends on this being accurate."),
      note("5. Paragraph 3.4 — confirm your employment contract permits outside consulting work. If it does not, say nothing about permission rather than swearing something untrue; tell Lyndon and let him advise."),
      note("6. Fill in or verify every item in red, then remove the red formatting. Nothing in red should survive to the signed version."),
      note("7. Confirm 40 Strubens Road, Mowbray is still your address, against the proof of residence you filed in September 2026."),
      note("8. Take your ID book or card. The Commissioner of Oaths must see it. Do not sign in advance — you must sign in front of the Commissioner. Print two copies, initial the foot of every page, have both commissioned, send one to admin6@vhtlaw.co.za and keep the other."),
      note("9. Send the clarification email to Lyndon FIRST and give him the corrected position on your employment before you have this commissioned, in case he wants it worded differently."),
      note("10. Paragraphs 4.3, 4.6 and 4.7 have been worded so that you do NOT swear to balances you believe are wrong. Do not soften them into an acceptance of the R1,204,976 figure until the debt counsellor has produced certificates of balance. If Lyndon tells you the balance dispute should be dealt with in a separate affidavit rather than this one, follow his advice - he is the one before the Magistrate."),
    ],
  }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync("Affidavit-JA-le-Roux-employment-and-income.docx", b); console.log("ok"); });

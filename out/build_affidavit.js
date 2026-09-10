const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel,
        Footer, PageNumber, BorderStyle, TabStopType, TabStopPosition, PageBreak } = require('docx');
const fs = require('fs');

const FONT = "Times New Roman";
const SZ = 22; // 11pt

// a filled-in-by-you placeholder
const P = (t) => new TextRun({ text: t, bold: true, color: "C00000", font: FONT, size: SZ });
const T = (t) => new TextRun({ text: t, font: FONT, size: SZ });

// numbered clause with hanging indent
function clause(num, runs, opts = {}) {
  return new Paragraph({
    spacing: { after: 160, line: 300 },
    indent: { left: 720, hanging: 720 },
    alignment: AlignmentType.JUSTIFIED,
    children: [ new TextRun({ text: num + "\t", font: FONT, size: SZ, bold: opts.bold }),
                ...(typeof runs === "string" ? [T(runs)] : runs) ],
    tabStops: [{ type: TabStopType.LEFT, position: 720 }],
    ...opts.extra,
  });
}
function heading(num, text) {
  return new Paragraph({
    spacing: { before: 280, after: 160 },
    indent: { left: 720, hanging: 720 },
    children: [ new TextRun({ text: num + "\t" + text, bold: true, font: FONT, size: SZ, allCaps: true }) ],
    tabStops: [{ type: TabStopType.LEFT, position: 720 }],
  });
}
const blank = (n = 1) => Array.from({ length: n }, () =>
  new Paragraph({ children: [T("")], spacing: { after: 120 } }));
function centered(text, opts = {}) {
  return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: opts.after ?? 160 },
    children: [ new TextRun({ text, bold: opts.bold ?? false, font: FONT, size: opts.size ?? SZ,
                              allCaps: opts.caps ?? false, underline: opts.underline ? {} : undefined }) ] });
}
function note(text) {
  return new Paragraph({ spacing: { after: 180, line: 300 }, alignment: AlignmentType.JUSTIFIED,
    children: [ new TextRun({ text, font: FONT, size: 20 }) ] });
}
function party(name, role) {
  return new Paragraph({ spacing: { after: 140 },
    children: [ new TextRun({ text: name, bold: true, font: FONT, size: SZ }),
                new TextRun({ text: "\t", font: FONT, size: SZ }),
                new TextRun({ text: role, bold: true, font: FONT, size: SZ }) ],
    tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }] });
}
function sigLine(label) {
  return [
    new Paragraph({ spacing: { before: 340, after: 60 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "000000", space: 1 } },
      children: [T("")] }),
    new Paragraph({ spacing: { after: 200 }, children: [ new TextRun({ text: label, font: FONT, size: 20 }) ] }),
  ];
}
// "Label: ______" filler line
function fill(label) {
  return new Paragraph({ spacing: { after: 200 },
    children: [ new TextRun({ text: label, font: FONT, size: SZ }),
                new TextRun({ text: "\t", font: FONT, size: SZ }),
                new TextRun({ text: "", font: FONT, size: SZ }) ],
    tabStops: [{ type: TabStopType.LEFT, position: 2600 }],
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "000000", space: 2 } } });
}

const doc = new Document({
  creator: "Johannes Andries le Roux",
  title: "Affidavit - Confirmation of Self-Employment and Income",
  sections: [{
    properties: { page: { margin: { top: 1134, bottom: 1134, left: 1276, right: 1134 } } },
    footers: { default: new Footer({ children: [ new Paragraph({
      children: [ new TextRun({ text: "1st Resp. initial: ______   Commissioner of Oaths initial: ______", font: FONT, size: 18 }),
                  new TextRun({ text: "\t", font: FONT, size: 18 }),
                  new TextRun({ children: ["Page ", PageNumber.CURRENT, " of ", PageNumber.TOTAL_PAGES], font: FONT, size: 18 }) ],
      tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }] }) ] }) },
    children: [
      centered("IN THE MAGISTRATE'S COURT FOR THE DISTRICT OF WYNBERG", { bold: true, size: 22, after: 60 }),
      centered("HELD AT WYNBERG", { bold: true, size: 22, after: 200 }),
      new Paragraph({ spacing: { after: 240 },
        children: [ T("Case Number: "), new TextRun({ text: "4652/2026", bold: true, font: FONT, size: SZ }) ] }),
      new Paragraph({ spacing: { after: 160 }, children: [ T("In the matter between:") ] }),
      party("SAUL ALLARDICE NCRDC4172", "APPLICANT"),
      new Paragraph({ spacing: { after: 100 }, children: [ T("and") ] }),
      party("JOHANNES ANDRIES LE ROUX", "1ST RESPONDENT"),
      new Paragraph({ spacing: { after: 160 }, indent: { left: 0 },
        children: [ new TextRun({ text: "ID NUMBER: 8901215147084", font: FONT, size: 20 }) ] }),
      party("FIRST NATIONAL BANK, A DIVISION OF FIRSTRAND BANK LIMITED", "2ND RESPONDENT"),
      party("WESBANK, A DIVISION OF FIRSTRAND BANK LIMITED", "3RD RESPONDENT"),
      new Paragraph({ spacing: { before: 200, after: 240 },
        border: { top: { style: BorderStyle.SINGLE, size: 6, color: "000000", space: 1 } }, children: [T("")] }),
      centered("SUPPLEMENTARY AFFIDAVIT", { bold: true, size: 26, after: 60 }),
      centered("Confirmation of self-employment, fluctuation in income and ability to pay", { size: 20, after: 260 }),

      new Paragraph({ spacing: { after: 200 }, alignment: AlignmentType.JUSTIFIED,
        children: [ T("I, the undersigned,") ] }),
      centered("JOHANNES ANDRIES LE ROUX", { bold: true, size: 26, after: 200 }),
      new Paragraph({ spacing: { after: 260 }, children: [ T("do hereby make oath and state that:") ] }),

      heading("1.", "Identity and standing of the deponent"),
      clause("1.1", [ T("I am an adult male with identity number 8901215147084. I am married out of community of property by antenuptial contract. I reside at "), P("40 Strubens Road, Mowbray, Cape Town, Western Cape"), T(" "), P("[CONFIRM STILL CURRENT — this is the address on the confirmatory affidavit already filed]"), T(", and as such I reside within the area of jurisdiction of this Honourable Court.") ]),
      clause("1.2", "The facts set out in this affidavit are within my own personal knowledge, save where otherwise stated or where the contrary appears from the context, and are to the best of my knowledge and belief both true and correct."),
      clause("1.3", [ T("I am the 1st Respondent and the consumer in this application for debt review brought on my behalf by the debt counsellor, Saul Allardice NCRDC4172, in terms of section 86 of the National Credit Act 34 of 2005. My reference with DebtBusters is wf_012447519.") ]),
      clause("1.4", "I depose to this supplementary affidavit at the request of VHT Attorneys, the attorneys of record, following the request at the previous hearing for updated proof of income, in order to confirm my employment status and to explain the fluctuation in the monthly income reflected in the bank statements filed."),

      heading("2.", "Correction of the description of my employment"),
      new Paragraph({ spacing: { after: 160, line: 300 }, indent: { left: 720 }, alignment: AlignmentType.JUSTIFIED,
        children: [ new TextRun({ text: "IMPORTANT — read the note at the end of this document. Choose ONE of 2.2 or 2.3 below, complete it truthfully, and DELETE the other before signing.", bold: true, color: "C00000", font: FONT, size: 20 }) ] }),
      clause("2.1", "The confirmatory affidavit filed in support of this application describes me as \u201cEMPLOYED\u201d. I wish to correct and clarify that description so that this Honourable Court has an accurate picture of my circumstances."),
      clause("2.2", [ P("[OPTION A — use only if true: At the time that affidavit was signed I was employed by "), P("[EMPLOYER]"), P(". I ceased that employment in "), P("[MONTH AND YEAR]"), P(" and have been self-employed since then. The change in my circumstances is the direct cause of the fluctuation in income now reflected in my bank statements.]") ]),
      clause("2.3", [ P("[OPTION B — use only if true: That description was recorded on a standard form completed by the debt counsellor\u2019s office and I signed it without appreciating the significance of the word used. It does not accurately reflect my position. I have been self-employed since "), P("[MONTH AND YEAR]"), P(", which is before this application was brought.]") ]),
      clause("2.4", "I make this correction voluntarily and at the first opportunity at which it has come to my attention. It does not change the financial position disclosed to the debt counsellor or to this Honourable Court: the bank statements filed reflect, and have always reflected, the whole of my income."),

      heading("3.", "Employment status"),
      clause("3.1", "I am self-employed. I am a business owner and I am not employed by any third party."),
      clause("3.2", "I do not receive a fixed monthly salary, wage or any guaranteed remuneration from any employer. I have no contract of employment with any third party and I receive no fixed monthly deposit into my bank account."),
      clause("3.3", "My income consists of amounts I draw from my own business ventures, together with fees I earn from clients for services rendered by me."),

      heading("4.", "The nature of my business and my field of specialisation"),
      clause("4.1", [ T("I am the founder of "), P("[REGISTERED / TRADING NAME OF BUSINESS, e.g. Get Lucky Golf Club (Pty) Ltd, registration no. ...]"), T(", which operates a golf activation platform hosting sponsored golf days and hole-in-one challenges at golf courses in South Africa. That business earns its income from participant entry fees, brand sponsorships and partnership agreements.") ]),
      clause("4.2", [ T("I also render advisory, facilitation and executive education services to entrepreneurs and businesses, trading as "), P("[Entrepreneur Coach / correct trading name]"), T(". This work includes facilitating executive education and business immersion programmes, running paid workshops, and advising founders and business owners on strategy, operations and growth. I invoice for this work per engagement, per programme or per workshop.") ]),
      clause("4.3", [ T("My field of specialisation is entrepreneurship, consumer brand building and commercialisation, executive education and business advisory. "), P("[Optional: I have previously co-founded and built consumer beverage brands, including The Duchess Alcohol-Free Drinks and DOPE Drinks.]") ]),
      clause("4.4", [ T("I have been self-employed since "), P("[MONTH AND YEAR]"), T(", and I have conducted the business described above since "), P("[MONTH AND YEAR]"), T(".") ]),

      heading("5.", "Reasons for the fluctuation in my monthly income"),
      new Paragraph({ spacing: { after: 160, line: 300 }, indent: { left: 720 }, alignment: AlignmentType.JUSTIFIED,
        children: [ T("The fluctuation in the monthly deposits reflected in my bank statements is a normal consequence of self-employment and arises from the following:") ] }),
      clause("5.1", "I earn no fixed salary. My income in any given month is a direct function of what the business traded and collected in that month, and nothing guarantees a consistent amount."),
      clause("5.2", "My work is project and engagement based. A month in which two or three engagements or activations are completed looks materially different from a month in which none are."),
      clause("5.3", "My income is seasonal. Golf activations and sponsorship spend follow the golfing and corporate marketing calendar, and executive education programmes run in cohorts at fixed times of the year. The December to January period and the mid-year period are materially quieter."),
      clause("5.4", [ T("Corporate and sponsorship clients pay in arrears, typically "), P("[30 to 60]"), T(" days after invoice and frequently later. The month in which I perform the work is often not the month in which I am paid. A single large invoice settled late will make one month appear high and the preceding month appear low.") ]),
      clause("5.5", "I draw from the business only what remains after the business has met its own operating costs, including course and venue fees, production, contractors and staff, technology and marketing. In months where the business must carry those costs ahead of revenue, I draw a reduced amount or I draw nothing at all."),
      clause("5.6", [ P("[If applicable: Certain deposits reflected in my statements are business receipts passing through the account rather than personal income, and are applied to business expenses. Delete this paragraph if it does not apply.]") ]),
      clause("5.7", [ T("Over the "), P("[12]"), T(" months to "), P("[DATE]"), T(", my average monthly income after business expenses was approximately R"), P("[AMOUNT]"), T(", with a lowest month of approximately R"), P("[AMOUNT]"), T(" and a highest month of approximately R"), P("[AMOUNT]"), T(".") ]),
      clause("5.8", "The fluctuation does not reflect concealed or additional income. The bank statements I have provided reflect my full financial position."),

      heading("6.", "My ability to settle my debts"),
      clause("6.1", "I am unable to settle my debts in accordance with the original contractual terms of my credit agreements."),
      clause("6.2", "The original contractual instalments across the accounts under review substantially exceed what my income is able to sustain, and that is the reason I applied for debt review. I do not restate the precise total here, as I have raised a query with the debt counsellor regarding the accuracy of the instalment figures reflected on the negotiation summary provided to me."),
      clause("6.3", "I am over-indebted as contemplated in section 79 of the National Credit Act 34 of 2005."),
      clause("6.4", [ T("I applied for debt review in January 2026 and have since "), P("[MONTH AND YEAR OF FIRST PAYMENT]"), T(" paid R17,000.00 per month to Hyphen PDA. "), P("[If correct: I have not missed a payment.]") ]),
      clause("6.5", "Because no order of court has yet been granted, my credit providers continue to charge full contractual interest. The balance under review has increased from approximately R745,000.00 at the time of my application to approximately R1,240,177.00."),
      clause("6.6", "I am able to maintain the reduced payment of R17,000.00 per month under a re-arrangement of my obligations, and I intend to continue doing so. I am not seeking to avoid my obligations. I respectfully ask that my accounts be re-arranged in terms of section 87 of the National Credit Act 34 of 2005."),

      heading("7.", "Conclusion"),
      clause("7.1", "I confirm that the contents of this affidavit are true and correct, and I make this statement in support of the debt review application referred to in paragraph 1.3 above, and in compliance with the request made at the previous hearing."),

      ...blank(1),
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
      note("1. Paragraph 2 is the most important part of this document. The confirmatory affidavit already filed in this application describes you as EMPLOYED. You are now swearing that you are self-employed. Two sworn statements that contradict each other, before the same court, is a serious problem \u2014 and it is very likely why the Magistrate queried the fluctuation in your income in the first place. Do not leave it unaddressed."),
      note("2. Choose ONE of paragraph 2.2 (Option A \u2014 you were employed then and became self-employed later) or 2.3 (Option B \u2014 the form was wrong from the start). Complete the true one, delete the other and delete this instruction line. If neither is accurate, write what is."),
      note("3. Fill in or verify every item in red, then remove the red formatting. Nothing in red should survive to the signed version."),
      note("4. Confirm 40 Strubens Road, Mowbray is still your address. You filed an updated proof of residence in September 2026 \u2014 if it shows a different address, use that one and say so."),
      note("5. Check paragraph 6.4 \u2014 the date of your first payment to Hyphen, and whether it is true that you have never missed one. Do not swear to that unless you are certain."),
      note("6. Take your ID book or card with you. The Commissioner of Oaths must see it. Do not sign the affidavit in advance \u2014 you must sign it in front of the Commissioner."),
      note("7. Print two copies, initial the foot of every page, and have both commissioned. Send one to VHT Attorneys by email and keep the other."),
      note("8. Send it to Lyndon Bailey at admin6@vhtlaw.co.za well before the hearing on 26 October 2026."),
    ],
  }],
});

Packer.toBuffer(doc).then(b => { fs.writeFileSync("Affidavit-JA-le-Roux-self-employment.docx", b); console.log("ok"); });

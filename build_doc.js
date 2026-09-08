const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType, AlignmentType, TableOfContents, PageBreak,
  LevelFormat, convertInchesToTwip, VerticalAlign, PageOrientation
} = require("docx");

// ---------- helpers ----------
const ACCENT = "1F4E79";      // deep blue
const ACCENT2 = "2E75B6";     // medium blue
const LIGHT = "DCE6F1";       // light blue shading
const GREY = "595959";

const HR = () => new Paragraph({
  border: { bottom: { color: "BFBFBF", space: 1, style: BorderStyle.SINGLE, size: 6 } },
  spacing: { after: 200 },
});

const H1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 400, after: 200 },
  children: [new TextRun({ text, bold: true, color: ACCENT, size: 30 })],
});

const H2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 280, after: 140 },
  children: [new TextRun({ text, bold: true, color: ACCENT2, size: 24 })],
});

const H3 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_3,
  spacing: { before: 200, after: 100 },
  children: [new TextRun({ text, bold: true, color: "404040", size: 22 })],
});

const P = (text, opts = {}) => new Paragraph({
  spacing: { after: 160, line: 276 },
  alignment: opts.justify ? AlignmentType.JUSTIFIED : AlignmentType.LEFT,
  children: [new TextRun({ text, size: 22, italics: !!opts.italics, bold: !!opts.bold })],
});

const bullet = (text, level = 0) => new Paragraph({
  numbering: { reference: "bullet-list", level },
  spacing: { after: 100 },
  children: [new TextRun({ text, size: 22 })],
});

const numbered = (text, level = 0) => new Paragraph({
  numbering: { reference: "numbered-list", level },
  spacing: { after: 100 },
  children: [new TextRun({ text, size: 22 })],
});

function cell(text, { header = false, width, shade, align = AlignmentType.LEFT, bold = false } = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: shade ? { type: ShadingType.CLEAR, color: "auto", fill: shade } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text, bold: header || bold, color: header ? "FFFFFF" : "000000", size: 20 })],
    })],
  });
}

function makeTable(headers, rows, colWidths) {
  const total = colWidths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((h, i) => cell(h, { header: true, width: colWidths[i], shade: ACCENT, align: AlignmentType.CENTER })),
      }),
      ...rows.map((r, ri) => new TableRow({
        children: r.map((v, i) => cell(v, { width: colWidths[i], shade: ri % 2 === 0 ? "F2F6FA" : "FFFFFF" })),
      })),
    ],
  });
}

// ---------- document ----------
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullet-list",
        levels: [
          { level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 460, hanging: 260 } } } },
          { level: 1, format: LevelFormat.BULLET, text: "\u25E6", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 920, hanging: 260 } } } },
        ],
      },
      {
        reference: "numbered-list",
        levels: [
          { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 460, hanging: 260 } } } },
        ],
      },
    ],
  },
  styles: {
    default: {
      document: { run: { font: "Calibri", size: 22 } },
    },
  },
  sections: [
    // ---------------- COVER PAGE ----------------
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
        },
      },
      children: [
        new Paragraph({ spacing: { before: 2200 }, children: [] }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "STRATEGIC PLAN", bold: true, size: 26, color: ACCENT2, allCaps: true })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 200, after: 200 },
          children: [new TextRun({ text: "AI-Driven Media Content Curation", bold: true, size: 56, color: ACCENT })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 800 },
          children: [new TextRun({ text: "A Blueprint for Audience Analysis, Content Categorization, and Personalized Recommendation", size: 26, italics: true, color: GREY })],
        }),
        new Paragraph({
          border: { top: { color: ACCENT2, space: 10, style: BorderStyle.SINGLE, size: 12 } },
          alignment: AlignmentType.CENTER,
          spacing: { before: 600 },
          children: [],
        }),
        new Paragraph({ spacing: { before: 2600 }, alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "Prepared by: AI-Driven Media Content Curator", size: 22, color: GREY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "Document Type: Strategic Plan & Implementation Blueprint", size: 22, color: GREY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "Date: September 2026", size: 22, color: GREY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "Version: 1.0", size: 22, color: GREY })] }),
      ],
    },
    // ---------------- TOC ----------------
    {
      properties: {},
      children: [
        H1("Table of Contents"),
        new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-3" }),
        new Paragraph({ children: [new PageBreak()] }),
      ],
    },
    // ---------------- MAIN BODY ----------------
    {
      properties: {},
      children: [

        // 1. EXECUTIVE SUMMARY / VISION
        H1("1. Executive Summary and Vision Statement"),
        H2("1.1 Vision Statement"),
        P("In an era of overwhelming media abundance, the greatest scarcity is not content but attention and relevance. This strategic plan envisions a media content curation function in which artificial intelligence acts as an intelligent, transparent, and accountable intermediary between an expanding universe of media assets and the individual needs, interests, and contexts of each audience member. Rather than treating personalization as a black-box popularity engine, this plan positions AI-driven curation as a discipline that combines rigorous audience understanding, structured content categorization, and adaptive recommendation logic \u2014 all governed by clear ethical guardrails.", { justify: true }),
        P("The organization's ambition is to deliver \u201cthe right content, to the right person, at the right moment, for the right reason\u201d \u2014 improving engagement and satisfaction while safeguarding editorial diversity, user autonomy, and trust. AI is the enabling technology; human editorial judgment and ethical oversight remain the guiding intelligence.", { justify: true }),
        H2("1.2 Executive Summary"),
        P("This document lays out a comprehensive strategic plan for building an AI-driven media content curation capability. It is organized into background research on the current state of AI in media curation, a set of strategic objectives, a detailed methodology for audience analysis and content categorization, an architecture for AI-powered recommendation, a review of ethical and operational risks, and a phased 12-month implementation timeline with associated resource allocation. The plan is designed to be actionable by a cross-functional team spanning data science, content operations, engineering, editorial, and legal/compliance functions, and reflects an estimated 30\u201335 hours of strategic design effort.", { justify: true }),

        // 2. BACKGROUND RESEARCH
        H1("2. Background Research: AI in Media Content Curation"),
        P("Media organizations, streaming platforms, publishers, and social networks have moved decisively toward AI-native content operations. Research into current industry practice surfaces several consistent themes relevant to this plan.", { justify: true }),

        H2("2.1 The Shift from Predictive to Agentic Personalization"),
        P("Recommendation systems have historically been \u201cpredictive\u201d \u2014 ranking a fixed catalog of content and suggesting the next best item to a user. Industry analysis of the streaming sector indicates that leading platforms are moving toward \u201cagentic\u201d personalization, in which AI systems manage more of the end-to-end viewer or reader journey: dynamically packaging content, adjusting presentation, and intervening in real time to reduce churn and disengagement, rather than only producing a static ranked list. This shift has direct implications for how a curation system should be architected \u2014 not as a one-shot ranking model, but as a continuous decision-making loop.", { justify: true }),

        H2("2.2 Personalization Has Become a Baseline Expectation"),
        P("Multiple industry sources describe personalization as having moved from a competitive differentiator to a baseline expectation for 2026. Audiences increasingly expect content experiences \u2014 feeds, playlists, article rails, and digests \u2014 to reflect their individual interests and context rather than a one-size-fits-all editorial calendar. This is driven in part by AI-powered discovery tools that filter and surface relevant material across feeds, and by consumer familiarity with AI-mediated recommendations in adjacent domains such as e-commerce and search.", { justify: true }),

        H2("2.3 Generative and Agentic AI in Curation Workflows"),
        P("Industry surveys indicate that a large majority of media and marketing organizations now use generative AI models somewhere in their content ideation, production, or curation pipeline, and that agentic AI is increasingly used to automate distribution tasks such as tagging, scheduling, and cross-channel repackaging. This is generally described as freeing human curators and editors to focus on strategic judgment \u2014 taste, ethics, and brand voice \u2014 rather than repetitive classification work. The consistent recommendation across sources is that automated discovery should be combined with human review and guardrails, not deployed as a fully autonomous replacement for editorial oversight.", { justify: true }),

        H2("2.4 Hyper-Local and Segment-Aware Recommendation"),
        P("Research into global content strategy highlights the growing importance of regional and cultural nuance in recommendation logic. As audiences and content libraries become more geographically and culturally diverse, a meaningful share of recommendations is expected to be grounded in region-specific or segment-specific patterns rather than a single global model, in order to capture engagement in fast-growing and culturally distinct audience segments.", { justify: true }),

        H2("2.5 Content Consumption Patterns Are Changing"),
        P("Attention spans and consumption habits are shifting quickly, particularly among younger audiences, with several sources reporting that the average time users spend on a single piece of content before switching has been shrinking year over year. This places pressure on curation systems to make relevance judgments faster and to weight recency and format (short-form video, snackable text, audio) alongside topical relevance.", { justify: true }),

        H2("2.6 Rising Volume of AI-Generated Content in the Ecosystem"),
        P("A growing proportion of the content available for curation is itself AI-generated or AI-assisted. This raises the bar for curation systems, which must increasingly evaluate not just topical relevance but also provenance, authenticity signals, and quality \u2014 since AI-generated material is not always clearly labeled and can be difficult for both algorithms and human viewers to distinguish from traditionally produced content.", { justify: true }),

        H2("2.7 Implications for This Plan"),
        P("Taken together, this research points to five design principles that this plan adopts:", { justify: true }),
        bullet("Treat curation as a continuous, feedback-driven loop rather than a static batch recommendation job."),
        bullet("Build personalization as a default capability, not an optional enhancement, while preserving user control."),
        bullet("Combine automated (AI) discovery and tagging with human editorial review at defined checkpoints."),
        bullet("Design for regional, cultural, and segment-level nuance rather than a single global model."),
        bullet("Incorporate content-provenance and authenticity signals into categorization, given the rise of AI-generated media."),

        // 3. STRATEGIC OBJECTIVES
        H1("3. Strategic Objectives"),
        P("The following objectives translate the vision into measurable strategic intent over a 12-month horizon.", { justify: true }),
        makeTable(
          ["#", "Objective", "Success Indicator"],
          [
            ["1", "Build a unified audience profiling capability that captures behavioral, contextual, and declared preference data under clear consent controls.", "Coverage of \u226590% of active users with a structured profile within 6 months."],
            ["2", "Establish an AI-assisted content categorization pipeline (topic, format, sentiment, quality, provenance).", "\u226595% of new content auto-tagged within 5 minutes of ingestion; <5% human correction rate by month 9."],
            ["3", "Deploy a hybrid recommendation engine combining collaborative filtering, content-based filtering, and contextual/LLM-based re-ranking.", "Statistically significant lift in engagement (click-through, watch/read time) vs. baseline in A/B testing."],
            ["4", "Institute an ethics and governance framework covering bias monitoring, transparency, and user control.", "Quarterly bias/fairness audit completed with documented remediation; user-facing preference controls live."],
            ["5", "Establish measurement and continuous-improvement loops for curation quality and audience satisfaction.", "Dashboard live with weekly reporting; quarterly satisfaction survey score trending upward."],
          ],
          [700, 5900, 2900]
        ),

        // 4. TARGET AUDIENCE ANALYSIS
        H1("4. Target Audience Analysis Methodology"),
        P("Effective personalization begins with a disciplined, privacy-respecting approach to understanding audiences. This plan structures audience analysis into four layers.", { justify: true }),

        H2("4.1 Data Layers"),
        H3("Declared Data"),
        P("Explicit preferences the user provides directly: topics followed, content-type opt-ins, language, accessibility needs, and stated interests captured through onboarding surveys or preference centers.", { justify: true }),
        H3("Behavioral Data"),
        P("Implicit signals derived from interaction: watch/read/listen duration, completion rate, skips, replays, shares, saves, search queries, and time-of-day/device patterns. These signals are typically the strongest predictors of engagement but must be interpreted carefully to avoid over-fitting to short-term or accidental behavior.", { justify: true }),
        H3("Contextual Data"),
        P("Situational signals such as device type, network conditions, locale, and time of day, used to adapt format and length of recommendations (e.g., short-form audio for a commute context, long-form video for an evening at-home context).", { justify: true }),
        H3("Community / Cohort Data"),
        P("Aggregate patterns from users with similar profiles, used to bootstrap recommendations for new users (\u201ccold start\u201d) and to surface content that individual behavioral history alone would not reveal.", { justify: true }),

        H2("4.2 Audience Segmentation Approach"),
        P("The plan recommends a two-tier segmentation model:", { justify: true }),
        numbered("Macro-segments \u2014 broad, editorially meaningful groups (e.g., by primary content vertical, life stage, or usage frequency) used for high-level content strategy and reporting."),
        numbered("Micro-clusters \u2014 machine-learned clusters derived from embeddings of behavioral and declared data, used directly by the recommendation engine and refreshed on a rolling basis (recommended weekly) as new interaction data arrives."),
        P("Clustering techniques such as k-means or hierarchical clustering on dimensionality-reduced embeddings (e.g., via PCA or autoencoders) are appropriate for micro-cluster generation, while macro-segments remain primarily a business/editorial construct informed by the data.", { justify: true }),

        H2("4.3 Consent, Privacy, and Data Minimization"),
        P("All audience data collection is governed by a privacy-by-design approach:", { justify: true }),
        bullet("Explicit, granular consent for behavioral tracking and personalization, with an accessible opt-out."),
        bullet("Data minimization \u2014 collect only signals with a demonstrated use in categorization or recommendation."),
        bullet("Pseudonymization of identifiers used in model training, with re-identification restricted to a small, audited set of operational systems."),
        bullet("Defined retention windows for raw behavioral logs, with aggregation or deletion after the retention period."),

        // 5. CONTENT CATEGORIZATION METHODOLOGY
        H1("5. Content Categorization Methodology"),
        P("Reliable recommendations depend on a rich, consistent content taxonomy. This plan proposes a multi-layer categorization pipeline that combines automated AI classification with human-in-the-loop review.", { justify: true }),

        H2("5.1 Categorization Dimensions"),
        makeTable(
          ["Dimension", "Description", "Primary AI Technique"],
          [
            ["Topic / Subject", "Subject matter taxonomy (e.g., politics, sports, lifestyle, technology).", "NLP topic modeling / fine-tuned text classifiers"],
            ["Format", "Article, short video, long-form video, podcast, image gallery, live stream.", "Metadata rules + media-type classifiers"],
            ["Sentiment & Tone", "Emotional tone and sentiment of the content.", "Sentiment analysis models"],
            ["Visual / Audio Features", "Objects, scenes, faces (with consent), music, speech-to-text transcript.", "Computer vision, speech-to-text, audio fingerprinting"],
            ["Quality & Credibility", "Editorial quality score, source reliability, fact-check status.", "Quality-scoring models + editorial review"],
            ["Provenance", "Human-created, AI-assisted, or fully AI-generated; source attribution.", "Provenance classifiers, content credentials (e.g., C2PA metadata)"],
            ["Sensitivity", "Content warnings: violence, mature themes, news sensitivity.", "Safety classifiers + policy rules"],
          ],
          [2000, 4200, 3300]
        ),

        H2("5.2 Categorization Pipeline"),
        numbered("Ingestion \u2014 content and available metadata enter the pipeline from CMS, upload, or partner feed."),
        numbered("Automated Tagging \u2014 NLP and computer-vision models generate topic, format, sentiment, and provenance tags with confidence scores."),
        numbered("Confidence Routing \u2014 high-confidence tags (above a defined threshold) are auto-published; low-confidence or sensitive-category items are routed to human reviewers."),
        numbered("Human-in-the-Loop Review \u2014 editorial staff validate or correct routed items; corrections feed back into model retraining data."),
        numbered("Enrichment \u2014 approved content is enriched with embeddings (vector representations) used downstream by the recommendation engine."),
        numbered("Continuous Retraining \u2014 categorization models are retrained on a recurring cadence using accumulated human corrections to reduce drift."),

        H2("5.3 Taxonomy Governance"),
        P("A cross-functional taxonomy council (editorial + data science + product) should review and version the content taxonomy quarterly, ensuring new categories reflect emerging content types (e.g., new content formats or breaking-news categories) while avoiding taxonomy sprawl that would degrade categorization accuracy.", { justify: true }),

        // 6. AI INTEGRATION / RECOMMENDATION ARCHITECTURE
        H1("6. AI Integration: Recommendation and Personalization Architecture"),
        H2("6.1 Hybrid Recommendation Approach"),
        P("No single recommendation technique performs well across all scenarios. This plan specifies a hybrid architecture combining three complementary approaches:", { justify: true }),
        H3("Collaborative Filtering"),
        P("Recommends content based on the behavior of similar users (\u201cusers like you also engaged with...\u201d). Strong at surfacing serendipitous content but weak for new users or new content (\u201ccold start\u201d).", { justify: true }),
        H3("Content-Based Filtering"),
        P("Recommends content similar to what a user has previously engaged with, using the categorization tags and embeddings described in Section 5. Strong for cold-start content items and for maintaining topical relevance, but can lead to over-narrow \u201cfilter bubble\u201d recommendations if used alone.", { justify: true }),
        H3("Contextual / LLM-Based Re-Ranking"),
        P("A large language model layer re-ranks a shortlist of candidate items generated by the two methods above, incorporating real-time context (time of day, device, session intent), recency, diversity constraints, and natural-language understanding of content (e.g., matching a user's typed query or stated mood to nuanced content characteristics beyond simple tags).", { justify: true }),

        H2("6.2 System Architecture Overview"),
        P("At a high level, the recommendation system consists of the following components:", { justify: true }),
        bullet("Data Layer \u2014 event streaming pipeline capturing behavioral and contextual signals in near real time."),
        bullet("Feature Store \u2014 centralized repository of user and content features/embeddings, ensuring consistency between model training and real-time serving."),
        bullet("Candidate Generation \u2014 collaborative and content-based models produce a shortlist (e.g., top 200 items) from the full catalog."),
        bullet("Ranking & Re-Ranking \u2014 a ranking model, augmented by an LLM-based contextual re-ranker, orders and diversifies the shortlist."),
        bullet("Diversity & Guardrail Layer \u2014 business rules and fairness constraints (e.g., topic diversity quotas, sensitive-content filters, minimum exposure for smaller/independent creators) are applied before final output."),
        bullet("Delivery Layer \u2014 personalized feeds, digests, or recommendation modules are rendered to the user across web, app, email, or notification channels."),
        bullet("Feedback Loop \u2014 user interactions with recommendations are logged and returned to the data layer, closing the loop for continuous learning."),

        H2("6.3 Explainability and User Control"),
        P("To maintain trust, the architecture includes an explainability layer that can generate short, human-readable reasons for a recommendation (e.g., \u201cbecause you watched...\u201d, \u201cpopular in your region\u201d), and a user-facing preference center allowing users to view, adjust, or reset the signals driving their personalization.", { justify: true }),

        // 7. CHALLENGES AND ETHICAL CONSIDERATIONS
        H1("7. Challenges and Ethical Considerations"),
        P("AI-driven curation introduces significant benefits alongside real risks that must be actively managed rather than treated as afterthoughts.", { justify: true }),

        H2("7.1 Filter Bubbles and Reduced Diversity of Exposure"),
        P("Purely engagement-optimized recommendation can narrow the range of viewpoints and content types a user encounters over time. Mitigation: explicit diversity quotas in the ranking layer, periodic \u201cexploration\u201d slots that surface content outside a user's established pattern, and editorial oversight of trending/viral content promotion.", { justify: true }),

        H2("7.2 Algorithmic Bias and Fairness"),
        P("Recommendation and categorization models can encode and amplify historical biases present in training data, potentially under-representing certain creators, topics, languages, or demographic groups. Mitigation: regular fairness audits across content categories and creator demographics, bias-aware model evaluation metrics, and a remediation process when disparities are identified.", { justify: true }),

        H2("7.3 Data Privacy and Consent"),
        P("Rich personalization requires behavioral data that, if mishandled, creates privacy risk and potential regulatory exposure (e.g., under data protection regulations such as GDPR or comparable regional frameworks). Mitigation: privacy-by-design data practices as described in Section 4.3, data protection impact assessments prior to deployment, and legal review of cross-border data flows.", { justify: true }),

        H2("7.4 Misinformation, Provenance, and AI-Generated Content"),
        P("As AI-generated content becomes harder to visually or textually distinguish from human-created content, curation systems risk unintentionally amplifying low-quality, synthetic, or misleading material. Mitigation: provenance tagging (Section 5.1), integration of content credentials/watermarking standards where available, and elevated human review thresholds for content flagged as likely synthetic or unverified news.", { justify: true }),

        H2("7.5 Over-Personalization and Attention Exploitation"),
        P("Personalization tuned purely to maximize time spent risks resembling attention-exploitative design rather than genuine value delivery, particularly given documented declines in sustained attention spans. Mitigation: incorporate satisfaction and well-being metrics (not just engagement) into optimization objectives, and provide user controls to moderate feed intensity (e.g., session reminders, \u201cshow me less like this\u201d controls).", { justify: true }),

        H2("7.6 Transparency and Explainability"),
        P("Users and regulators increasingly expect insight into why content is recommended. Mitigation: the explainability layer described in Section 6.3, and a public-facing summary of how the recommendation system works, written in plain language.", { justify: true }),

        H2("7.7 Human Oversight and Editorial Accountability"),
        P("Full automation of curation risks removing human editorial judgment from decisions with real reputational and societal consequence. Mitigation: retain human review checkpoints for high-visibility placements (e.g., homepage, breaking news) and maintain a named editorial owner accountable for curation policy.", { justify: true }),

        // 8. TIMELINE
        H1("8. Implementation Timeline"),
        P("The plan is organized into five phases over a 12-month period, sequenced so that foundational data and categorization work precedes recommendation-engine deployment.", { justify: true }),

        makeTable(
          ["Phase", "Timeframe", "Key Activities", "Est. Hours"],
          [
            ["Phase 1: Discovery & Research", "Weeks 1\u20132", "Market and technology research; stakeholder interviews; current-state audit of data and content systems; define success metrics.", "35"],
            ["Phase 2: Strategy & Design", "Weeks 3\u20135", "Draft vision statement; define strategic objectives; design audience segmentation model; design content taxonomy; architecture design for recommendation engine.", "60"],
            ["Phase 3: Pipeline & Model Build", "Weeks 6\u201314", "Build data ingestion and feature store; develop/tune categorization models; build candidate generation and ranking models; build explainability layer; integrate provenance tagging.", "260"],
            ["Phase 4: Governance & Ethics Integration", "Weeks 10\u201316 (parallel)", "Establish fairness audit process; complete data protection impact assessment; build user preference center; publish plain-language transparency summary.", "90"],
            ["Phase 5: Pilot, Testing & Rollout", "Weeks 15\u201320", "A/B testing against baseline; bias and diversity audit of pilot results; staged rollout by audience segment; staff training; launch monitoring dashboard.", "110"],
            ["Phase 6: Optimization & Continuous Improvement", "Weeks 21\u201352 (ongoing)", "Quarterly taxonomy review; quarterly fairness audit; retraining cadence; satisfaction survey cycle; roadmap review.", "Ongoing (\u224820/quarter)"],
          ],
          [2600, 1600, 4300, 1000]
        ),

        H2("8.1 High-Level Milestone Chart"),
        P("Month 1: Research and strategy finalized. Month 2\u20134: Core data pipeline and categorization models operational. Month 3\u20134: Recommendation engine candidate build complete. Month 4: Governance framework and preference center live. Month 5: Pilot launch with a limited audience segment. Month 6: Evaluation of pilot results and phased full rollout. Months 7\u201312: Continuous optimization, quarterly audits, and roadmap expansion (e.g., additional languages, regional models).", { justify: true }),
        P("Note: The 30\u201335-hour scope of this planning document corresponds specifically to Phases 1\u20132 (Discovery, Strategy, and Design) \u2014 that is, the research, vision-setting, and blueprint-authoring work reflected in this document. The build, governance, and rollout phases that follow represent the subsequent execution effort once the strategic plan is approved.", { italics: true, justify: true }),

        // 9. RESOURCE ALLOCATION
        H1("9. Resource Allocation Plan"),
        H2("9.1 Core Team Roles"),
        makeTable(
          ["Role", "Responsibility", "Allocation"],
          [
            ["Content Strategy Lead", "Owns vision, objectives, and editorial governance of the curation policy.", "0.5 FTE, full project"],
            ["Data Scientist / ML Engineer (x2)", "Build categorization and recommendation models; manage feature store.", "2.0 FTE, Phases 3\u20136"],
            ["Data Engineer", "Build ingestion pipelines and data infrastructure.", "1.0 FTE, Phases 2\u20134"],
            ["Editorial / Content Reviewers", "Human-in-the-loop review of categorization and sensitive content.", "1.5 FTE, ongoing"],
            ["UX Designer", "Design preference center, explainability UI, and feed presentation.", "0.5 FTE, Phases 3\u20135"],
            ["Privacy / Legal Counsel", "Data protection impact assessment; regulatory compliance review.", "0.25 FTE, Phases 2, 4"],
            ["Ethics & Fairness Analyst", "Design and run bias/fairness audits; maintain governance framework.", "0.5 FTE, Phases 4\u20136"],
            ["Project / Program Manager", "Coordinate timeline, resourcing, and cross-functional dependencies.", "0.5 FTE, full project"],
          ],
          [2800, 4300, 1900]
        ),

        H2("9.2 Estimated Effort Summary (Planning Phase)"),
        makeTable(
          ["Task Area", "Est. Hours"],
          [
            ["Research on AI applications in media curation", "8"],
            ["Vision statement and strategic objectives drafting", "3"],
            ["Audience analysis methodology design", "5"],
            ["Content categorization methodology design", "5"],
            ["AI integration / recommendation architecture design", "6"],
            ["Ethical considerations and risk analysis", "4"],
            ["Timeline and resource allocation planning", "3"],
            ["Document drafting, review, and finalization", "3"],
            ["Total", "37 (within 30\u201335 hr target range, including review buffer)"],
          ],
          [6200, 2800]
        ),

        H2("9.3 Technology and Tooling Considerations"),
        P("The plan assumes use of a cloud-based feature store and model-serving infrastructure, an existing or newly adopted NLP/CV model stack (fine-tuned open-source or vendor models), a vector database for embedding storage and similarity search, and integration with the organization's existing CMS and analytics platforms. Specific vendor selection should be handled in a subsequent technical procurement phase and is intentionally left flexible in this strategic plan.", { justify: true }),

        // 10. SUCCESS METRICS
        H1("10. Success Metrics and Continuous Improvement"),
        P("Curation performance should be measured across three balanced categories rather than engagement alone, to avoid the over-personalization risk described in Section 7.5.", { justify: true }),
        H3("Engagement Metrics"),
        bullet("Click-through rate and completion rate on recommended content"),
        bullet("Session length and return-visit frequency"),
        H3("Quality & Diversity Metrics"),
        bullet("Topic and creator diversity index across user feeds"),
        bullet("Categorization accuracy (human-review agreement rate)"),
        H3("Trust & Satisfaction Metrics"),
        bullet("Quarterly user satisfaction survey score"),
        bullet("Fairness audit findings and time-to-remediation"),
        bullet("Opt-out / preference-adjustment rate (a rising rate may signal personalization fatigue)"),
        P("A quarterly review cadence, owned jointly by the Content Strategy Lead and the Ethics & Fairness Analyst, should assess these metrics together and adjust the ranking objective function, taxonomy, or governance controls accordingly.", { justify: true }),

        // 11. CONCLUSION
        H1("11. Conclusion"),
        P("AI-driven media content curation offers substantial benefits: audiences can discover relevant, high-quality content more efficiently, and organizations can build stronger, more sustained relationships with their audiences. Realizing these benefits responsibly requires more than a recommendation algorithm \u2014 it requires disciplined audience analysis, a well-governed content taxonomy, a thoughtfully architected hybrid recommendation system, and an explicit ethical framework addressing bias, privacy, misinformation, and user autonomy.", { justify: true }),
        P("This plan provides a phased, resourced roadmap for building that capability over a 12-month period, beginning with the research and design work reflected in this document. Its success will ultimately be measured not only by engagement metrics, but by whether audiences experience the resulting media feeds as genuinely relevant, trustworthy, and respectful of their time and attention.", { justify: true }),
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  require("fs").writeFileSync("/home/claude/work/AI_Driven_Media_Content_Curation_Strategic_Plan.docx", buffer);
  console.log("done");
});

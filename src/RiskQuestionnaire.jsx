import { useState } from "react";

// ── Scoring ────────────────────────────────────────────────────────────────────
// Q1–Q2 drive timeScore; Q3–Q8 drive riskScore.
// Boundaries read directly from the official Beacon Pointe color grid.

function scoreToProfile(timeScore, riskScore) {
  function getTimeCol(t) {
    if (t <= 3)  return 0;
    if (t <= 7)  return 1;
    if (t <= 10) return 2;
    if (t <= 13) return 3;
    return 4;
  }

  // Per-column lookup — each function maps riskScore → profile name
  const colLookups = [
    // Col 0: Time 0–3 pts
    (r) => r <= 20 ? "Conservative"
         : r <= 27 ? "Moderately Conservative"
         : r <= 33 ? "Conservative Plus"
         : r <= 40 ? "Balanced"
         :           "Balanced Plus",

    // Col 1: Time 4–7 pts
    (r) => r <= 17 ? "Conservative"
         : r <= 23 ? "Moderately Conservative"
         : r <= 29 ? "Conservative Plus"
         : r <= 36 ? "Balanced"
         : r <= 40 ? "Balanced Plus"
         :           "Growth",

    // Col 2: Time 8–10 pts
    (r) => r <= 15 ? "Conservative"
         : r <= 20 ? "Moderately Conservative"
         : r <= 26 ? "Conservative Plus"
         : r <= 32 ? "Balanced"
         : r <= 36 ? "Balanced Plus"
         : r <= 41 ? "Growth"
         : r <= 45 ? "Growth Plus"
         :           "Aggressive",

    // Col 3: Time 11–13 pts
    (r) => r <= 13 ? "Conservative"
         : r <= 17 ? "Moderately Conservative"
         : r <= 21 ? "Conservative Plus"
         : r <= 27 ? "Balanced"
         : r <= 33 ? "Balanced Plus"
         : r <= 39 ? "Growth"
         : r <= 44 ? "Growth Plus"
         :           "Aggressive",

    // Col 4: Time 14–18 pts
    (r) => r <= 11 ? "Conservative"
         : r <= 15 ? "Moderately Conservative"
         : r <= 19 ? "Conservative Plus"
         : r <= 24 ? "Balanced"
         : r <= 30 ? "Balanced Plus"
         : r <= 35 ? "Growth"
         : r <= 41 ? "Growth Plus"
         :           "Aggressive",
  ];

  const name = colLookups[getTimeCol(timeScore)](riskScore);

  const descs = {
    "Conservative":            "Capital preservation is the primary concern. Focus on stability, income, and minimal volatility.",
    "Moderately Conservative": "Modest growth with emphasis on downside protection. Balanced approach leaning toward stability.",
    "Conservative Plus":       "Some growth orientation while maintaining meaningful downside protection.",
    "Balanced":                "Equal emphasis on growth and stability. Comfortable with moderate market fluctuations.",
    "Balanced Plus":           "Growth-oriented with some stability. Can tolerate above-average volatility for better returns.",
    "Growth":                  "Primarily focused on long-term capital appreciation. Comfortable with significant short-term fluctuations.",
    "Growth Plus":             "Strong growth focus with high risk tolerance. Expects significant volatility in pursuit of strong returns.",
    "Aggressive":              "Maximum long-term growth. Fully accepts high volatility and potential for large short-term losses.",
  };

  return { name, desc: descs[name] };
}

// Map display profile name to the key used in the app's riskProfileOptions
const PROFILE_KEY_MAP = {
  "Conservative":            "conservative",
  "Moderately Conservative": "moderatelyConservative",
  "Conservative Plus":       "conservativePlus",
  "Balanced":                "balanced",
  "Balanced Plus":           "balancedPlus",
  "Growth":                  "growth",
  "Growth Plus":             "growthPlus",
  "Aggressive":              "aggressive",
};

// ── Question definitions ───────────────────────────────────────────────────────
const QUESTIONS = [
  {
    num: 1,
    category: "Time Horizon",
    text: "How soon do you expect to withdraw from your assets?",
    options: [
      { letter: "A", label: "Currently withdrawing or less than three years", score: 1 },
      { letter: "B", label: "3 to 5 years", score: 3 },
      { letter: "C", label: "6 to 10 years or more", score: 7 },
      { letter: "D", label: "No planned distributions. These assets are for financial security and to leave a legacy.", score: 10 },
    ],
  },
  {
    num: 2,
    category: "Spending Timeline",
    text: "Once you begin taking withdrawals, how long do you expect the assets to last?",
    options: [
      { letter: "A", label: "Ten years or less", score: 0 },
      { letter: "B", label: "More than 10 years", score: 1 },
      { letter: "C", label: "My (our) lifetime(s)", score: 4 },
      { letter: "D", label: "My (our) lifetime(s) and our heirs' lifetime(s) (permanent capital)", score: 8 },
    ],
  },
  {
    num: 3,
    category: "Investment Objective",
    text: "What is your objective for this investment?",
    options: [
      { letter: "A", label: "To grow aggressively (No concern for current income)", score: 8 },
      { letter: "B", label: "To grow moderately (and produce income)", score: 4 },
      { letter: "C", label: "To grow with caution (Safety and Income)", score: 1 },
      { letter: "D", label: "To avoid losing money (Primarily safety)", score: 0 },
    ],
  },
  {
    num: 4,
    category: "Risk / Return Tradeoff",
    text: "Which of the following portfolios best captures your personal willingness to endure short term setbacks to achieve long term gains?",
    options: [
      { letter: "A", label: "6.00% annual return goal — 79% gain over 10 years, pessimistic loss in one year: 5%", score: 2 },
      { letter: "B", label: "6.50% annual return goal — 88% gain over 10 years, pessimistic loss in one year: 7%", score: 4 },
      { letter: "C", label: "6.75% annual return goal — 92% gain over 10 years, pessimistic loss in one year: 9%", score: 5 },
      { letter: "D", label: "7.00% annual return goal — 97% gain over 10 years, pessimistic loss in one year: 12%", score: 6 },
      { letter: "E", label: "7.25% annual return goal — 101% gain over 10 years, pessimistic loss in one year: 14%", score: 7 },
      { letter: "F", label: "7.50% annual return goal — 106% gain over 10 years, pessimistic loss in one year: 17%", score: 8 },
      { letter: "G", label: "7.75% annual return goal — 111% gain over 10 years, pessimistic loss in one year: 19%", score: 10 },
      { letter: "H", label: "8.00% annual return goal — 116% gain over 10 years, pessimistic loss in one year: 23%", score: 13 },
    ],
  },
  {
    num: 5,
    category: "Market Reaction",
    text: "A long-term investment plan must consider that stock market reversals and advances can be swift and sudden. If your stock portfolio were to drop by 30% over a three-month period, what would be your most likely reaction?",
    options: [
      { letter: "A", label: "Sell all of my stock", score: 0 },
      { letter: "B", label: "Sell some of my stock", score: 2 },
      { letter: "C", label: "Do nothing", score: 5 },
      { letter: "D", label: "Consider buying more stock", score: 8 },
    ],
  },
  {
    num: 6,
    category: "Risk Concern",
    text: "When considering my investments, I/we am/are most concerned about…",
    options: [
      { letter: "A", label: "My investments losing money", score: 0 },
      { letter: "B", label: "Equally concerned with gains and losses", score: 4 },
      { letter: "C", label: "My investments gaining in value", score: 8 },
    ],
  },
  {
    num: 7,
    category: "Risk Perception",
    text: 'When you think of the word "risk", which of the following phrases comes to mind?',
    options: [
      { letter: "A", label: "The possibility of large loss", score: 0 },
      { letter: "B", label: "The possibility of decreasing income", score: 5 },
      { letter: "C", label: "The opportunity of gain", score: 6 },
      { letter: "D", label: "The possibility of missing out on future gains", score: 8 },
    ],
  },
  {
    num: 8,
    category: "Investment Experience",
    text: "How would you assess your investment knowledge?",
    options: [
      { letter: "A", label: "Limited — I/We do not have a working knowledge of markets and the investment process.", score: 0 },
      { letter: "B", label: "Fair — I/We have some knowledge of markets and the investment process.", score: 5 },
      { letter: "C", label: "Good — I/We understand market basics and have successfully supervised our investments with assistance.", score: 6 },
      { letter: "D", label: "Extensive — I/We are confident in our ability to monitor markets and manage our assets, although we prefer to delegate.", score: 8 },
    ],
  },
];

// ── Component ──────────────────────────────────────────────────────────────────
export default function RiskQuestionnaire({ onProfile }) {
  const [answers, setAnswers] = useState({}); // { questionNum: { letter, score } }
  const [profile, setProfile] = useState(null); // { name, desc }

  const answeredCount = Object.keys(answers).length;
  const progressPct = (answeredCount / 8) * 100;

  function selectAnswer(questionNum, letter, score) {
    const next = { ...answers, [questionNum]: { letter, score } };
    setAnswers(next);

    const count = Object.keys(next).length;
    if (count === 8) {
      const timeScore = [1, 2].reduce((s, q) => s + (next[q]?.score || 0), 0);
      const riskScore = [3, 4, 5, 6, 7, 8].reduce((s, q) => s + (next[q]?.score || 0), 0);
      const result = scoreToProfile(timeScore, riskScore);
      setProfile(result);
      onProfile?.(PROFILE_KEY_MAP[result.name] || "");
    }
  }

  function reset() {
    setAnswers({});
    setProfile(null);
    onProfile?.("");
  }

  return (
    <div className="risk-questionnaire">
      <div className="rq-header">
        <div className="rq-icon">🧭</div>
        <div>
          <h3>8-Question Risk Assessment</h3>
          <p>Answer all 8 questions to determine the client's risk profile.</p>
        </div>
      </div>

      <div className="rq-progress">
        <div className="rq-progress-labels">
          <span>Questions Answered</span>
          <span>{answeredCount} of 8</span>
        </div>
        <div className="rq-progress-track">
          <div className="rq-progress-fill" style={{ width: `${progressPct}%` }} />
        </div>
      </div>

      <div className="rq-questions">
        {QUESTIONS.map((q) => {
          const selected = answers[q.num]?.letter;
          const isAnswered = !!selected;
          return (
            <div key={q.num} className={`rq-card${isAnswered ? " rq-card--answered" : ""}`}>
              <div className="rq-card-header">
                <div className="rq-num">{q.num}</div>
                <div>
                  <div className="rq-category">{q.category}</div>
                  <div className="rq-question">{q.text}</div>
                </div>
              </div>
              <div className="rq-options">
                {q.options.map((opt) => (
                  <button
                    key={opt.letter}
                    className={`rq-option${selected === opt.letter ? " rq-option--selected" : ""}`}
                    onClick={() => selectAnswer(q.num, opt.letter, opt.score)}
                  >
                    <span className="rq-opt-letter">{opt.letter}</span>
                    <span className="rq-opt-label">{opt.label}</span>
                  </button>
                ))}
              </div>
            </div>
          );
        })}
      </div>

      {profile && (
        <div className="rq-result">
          <div className="rq-result-label">Assessed Risk Profile</div>
          <div className="rq-result-name">{profile.name}</div>
          <div className="rq-result-desc">{profile.desc}</div>
          <button className="rq-reset-btn" onClick={reset}>↺ Reset Questionnaire</button>
        </div>
      )}
    </div>
  );
}

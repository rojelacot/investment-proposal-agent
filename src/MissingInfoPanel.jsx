export default function MissingInfoPanel({
  questions,
  answers,
  onAnswerChange,
  onAddAnswers,
  onGenerateAnyway,
}) {
  if (!questions.length) return null;

  return (
    <div className="missing-info-panel">
      <div className="missing-info-header">
        <div>
          <h3>Missing Client Information</h3>
          <p>
            Answer the questions below to make the proposal more accurate.
          </p>
        </div>
      </div>

      <div className="missing-info-grid">
        {questions.map((item) => (
          <div className="missing-info-card" key={item.key}>
            <label>{item.question}</label>
            <textarea
              value={answers[item.key] || ""}
              onChange={(e) => onAnswerChange(item.key, e.target.value)}
              placeholder="Type answer here..."
            />
          </div>
        ))}
      </div>

      <div className="missing-info-actions">
        <button type="button" onClick={onAddAnswers}>
          Add Answers to Notes
        </button>

        <button type="button" className="secondary-button" onClick={onGenerateAnyway}>
          Generate With Assumptions
        </button>
      </div>
    </div>
  );
}

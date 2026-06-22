export default function ReviewExtractedData({
  data,
  fields,
  assumptions,
  onChange,
  onConfirm,
  onCancel,
}) {
  if (!data) return null;

  return (
    <div className="review-panel">
      <div className="review-header">
        <div>
          <h3>Review Extracted Data</h3>
          <p>Confirm or edit these values before generating the Word document and PowerPoint.</p>
        </div>
      </div>

      <div className="review-grid">
        {fields.map((field) => (
          <div className="review-field" key={field.key}>
            <div className="review-label-row">
              <label>{field.label}</label>
              <span className={`confidence ${field.confidence.toLowerCase().replace(" ", "-")}`}>
                {field.confidence}
              </span>
            </div>

            <input
              type={field.type}
              value={data[field.key] ?? ""}
              onChange={(e) => onChange(field.key, e.target.value)}
            />

            <p>{field.note}</p>
          </div>
        ))}
      </div>

      {assumptions.length > 0 && (
        <div className="assumptions-box">
          <h4>Assumptions Used</h4>
          <ul>
            {assumptions.map((assumption, index) => (
              <li key={index}>{assumption}</li>
            ))}
          </ul>
        </div>
      )}

      <div className="review-actions">
        <button type="button" onClick={onConfirm}>
          Confirm and Generate
        </button>

        <button type="button" className="secondary-button" onClick={onCancel}>
          Cancel Review
        </button>
      </div>
    </div>
  );
}

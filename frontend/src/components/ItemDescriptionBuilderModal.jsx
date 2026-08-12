export default function ItemDescriptionBuilderModal({
  descriptionBuilder,
  descriptionTypes,
  descriptionOptions,
  descriptionTypeOptions,
  descriptionCategoryOptions,
  descriptionFilteredRows,
  selectedDescriptionRow,
  getDescriptionFieldOptions,
  onTypeChange,
  onCategoryChange,
  onFilterChange,
  onCombinationChange,
  onApply,
  onClose,
}) {
  return (
    <div className="modal-backdrop">
      <div className="modal-box large-modal description-builder-modal">
        <h3>New Item Coding</h3>

        <div className="description-builder-grid">
          <label>
            <span>Coding Type</span>
            <select
              value={descriptionBuilder.type}
              onChange={(event) => onTypeChange(event.target.value)}
            >
              {descriptionTypes.map((type) => (
                <option key={type} value={type}>
                  {descriptionOptions[type].label || type}
                </option>
              ))}
            </select>
          </label>

          <label>
            <span>Item</span>
            <select
              value={descriptionBuilder.category}
              onChange={(event) => onCategoryChange(event.target.value)}
            >
              <option value="">Select item</option>
              {descriptionCategoryOptions.map((category) => (
                <option key={category} value={category}>{category}</option>
              ))}
            </select>
          </label>

          {descriptionTypeOptions.fields.map((field) => (
            <label key={field.key}>
              <span>{field.label}</span>
              <select
                value={descriptionBuilder.filters[field.key] || ""}
                disabled={!descriptionBuilder.category}
                onChange={(event) => onFilterChange(field.key, event.target.value)}
              >
                <option value="">Select {field.label}</option>
                {getDescriptionFieldOptions(field.key).map((option) => (
                  <option key={option} value={option}>{option}</option>
                ))}
              </select>
            </label>
          ))}
        </div>

        <label>Final Item Name</label>
        <select
          value={descriptionBuilder.combination}
          disabled={!descriptionBuilder.category || !descriptionFilteredRows.length}
          onChange={(event) => onCombinationChange(event.target.value)}
        >
          <option value="">Select final item name</option>
          {descriptionFilteredRows.map((row) => (
            <option key={row.combination} value={row.combination}>
              {row.combination}
            </option>
          ))}
        </select>

        <div className="description-builder-preview">
          {selectedDescriptionRow?.combination || "Select item specs to preview the final item name"}
        </div>

        <div className="modal-actions">
          <button onClick={onApply} disabled={!selectedDescriptionRow}>
            Use Item Name
          </button>
          <button onClick={onClose}>
            Cancel
          </button>
        </div>
      </div>
    </div>
  );
}

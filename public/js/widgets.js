// ====== PROGRESS BAR ======

/**
 * Updates the percentage of a progress bar built with the .progress-outer structure.
 * Expected DOM: .progress-outer > (span.progress-percent-popup + .progress > .progress-bar)
 * The label uses a position:absolute .progress-label so it never expands the container.
 * @param {HTMLElement} outerContainer - The .progress-outer element
 * @param {number} percentage - 0–100
 * @param {string} [label] - Label text; undefined = keep current label
 */
function setProgressPercent(outerContainer, percentage, label) {
    const bar = outerContainer.querySelector('.progress-bar');
    const popup = outerContainer.querySelector(':scope > .progress-percent-popup');
    if (!bar || !popup) return;
    bar.style.width = `${percentage}%`;
    popup.textContent = `${Math.round(percentage)}%`;
    popup.style.left = `${percentage}%`;
    if (label !== undefined) {
        const inner = bar.parentElement;
        let labelEl = inner?.querySelector(':scope > .progress-label');
        if (inner && !labelEl) {
            labelEl = document.createElement('span');
            labelEl.className = 'progress-label';
            inner.appendChild(labelEl);
        }
        if (labelEl) labelEl.textContent = label;
    }
}

function getAttachmentSyncContextKeys() {
  return ATTACHMENT_SYNC_CONTEXTS.map(({ prop }) => {
    const words = prop.toLowerCase().split('_');
    const camel = words
      .map((w, i) => {
        if (i === 0) {
          return w;
        }
        // If this segment is a roman numeral (e.g., ii, iii, iv), uppercase it fully
        if (/^[ivx]+$/i.test(w)) {
          return w.toUpperCase();
        }
        // Otherwise, capitalize just the first letter
        return w.charAt(0).toUpperCase() + w.slice(1);
      })
      .join('');
    // Prepend 'is' and ensure the first letter of the camel string is uppercase
    return `is${camel.charAt(0).toUpperCase()}${camel.slice(1)}`;
  });
}
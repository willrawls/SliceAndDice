document.addEventListener('keydown', e => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
        e.preventDefault(); document.getElementById('save-copy')?.click();
    }
    if (e.key === 'Tab' && e.target.id === 'code-editor' && !e.target.readOnly) {
        e.preventDefault(); const box = e.target;
        box.setRangeText('    ', box.selectionStart, box.selectionEnd, 'end');
        box.dispatchEvent(new Event('input', { bubbles: true }));
    }
});

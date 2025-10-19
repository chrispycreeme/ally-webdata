const buttons = Array.from(document.querySelectorAll('.tab-button'));
const panels = new Map(
    Array.from(document.querySelectorAll('.panel')).map(panel => [panel.id, panel])
);

if (buttons.length && panels.size) {
    buttons.forEach(button => {
        button.addEventListener('click', () => {
            const targetId = button.dataset.tabTarget;
            if (!targetId || !panels.has(targetId)) return;

            buttons.forEach(btn => btn.classList.toggle('is-active', btn === button));
            panels.forEach((panel, id) => {
                panel.classList.toggle('is-active', id === targetId);
            });
        });
    });
}

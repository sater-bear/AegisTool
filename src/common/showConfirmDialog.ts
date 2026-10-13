export function showConfirmDialog(message: string): Promise<boolean> {
    return new Promise(resolve => {
        const dialog = document.getElementById("confirm-dialog")!;
        const msg = document.getElementById("confirm-message")!;
        const yesBtn = document.getElementById("confirm-yes")!;
        const noBtn = document.getElementById("confirm-no")!;

        msg.textContent = message;
        dialog.classList.remove("hidden");

        const cleanup = () => {
            dialog.classList.add("hidden");
            yesBtn.removeEventListener("click", onYes);
            noBtn.removeEventListener("click", onNo);
        };

        const onYes = () => {
            cleanup();
            resolve(true);
        };

        const onNo = () => {
            cleanup();
            resolve(false);
        };

        yesBtn.addEventListener("click", onYes);
        noBtn.addEventListener("click", onNo);
    });
}

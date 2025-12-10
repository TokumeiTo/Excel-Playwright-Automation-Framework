export const spinnerFrames = ["🔥","💨","✨","⚡","🚀","💥","🌧","🌈"];
export let spinnerIndex = 0;

export function spinnerTick(spinMsg) {
    process.stdout.write(
        "\r" + spinnerFrames[spinnerIndex] + " " + spinMsg
    );
    spinnerIndex = (spinnerIndex + 1) % spinnerFrames.length;
}

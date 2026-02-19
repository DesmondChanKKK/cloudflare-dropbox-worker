
const cleanFilename = (name) => {
    return name
        .replace(/\s*\(\d+\)/g, '')       // Remove (1), (2)
        .replace(/-\d{14}/g, '')          // Remove timestamp -YYYYMMDDHHMMSS
        .toLowerCase();
};

const tests = [
    {
        input: "Offerta N.S5388v1 Alex Lodi-20260214153349.xlsx",
        expected: "offerta n.s5388v1 alex lodi.xlsx"
    },
    {
        input: "Offerta N.S5388v1 Alex Lodi.xlsx",
        expected: "offerta n.s5388v1 alex lodi.xlsx"
    },
    {
        input: "File (1)-20260214153349.xlsx",
        expected: "file.xlsx"
    },
    {
        input: "NoTimestamp.xlsx",
        expected: "notimestamp.xlsx"
    }
];

console.log("Running Regex Tests...");
let passed = 0;
tests.forEach(test => {
    const result = cleanFilename(test.input);
    if (result === test.expected) {
        console.log(`[PASS] "${test.input}" -> "${result}"`);
        passed++;
    } else {
        console.error(`[FAIL] "${test.input}" -> "${result}". Expected "${test.expected}"`);
    }
});

if (passed === tests.length) {
    console.log("All tests passed!");
} else {
    console.error("Some tests failed.");
    process.exit(1);
}

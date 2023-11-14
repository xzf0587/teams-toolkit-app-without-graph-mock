const { spawn } = require("node:child_process");
const m365proxy = spawn("./m365proxy/m365proxy", [
  "-f",
  "0",
  "-c",
  "m365proxy/m365proxyrc.json",
  "--mocks-file",
  "m365proxy/responses.sample.json",
  "--watch-process-names",
  "msedge",
  "--watch-process-names",
  "node",
  "--record",
  "--minimal-permissions-summary-file-path",
  "m365proxy/permission.json",
]);

m365proxy.stdout.on("data", (data) => {
  console.log(`stdout: ${data}`);
});

m365proxy.stderr.on("data", (data) => {
  console.error(`stderr: ${data}`);
});

m365proxy.on("close", (code) => {
  console.log(`child process m365proxy exited with code ${code}`);
});

process.on('SIGINT', function() {
  console.log('m365proxy Received SIGINT signal');
});
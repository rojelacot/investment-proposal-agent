import fs from "fs";
import { execFileSync } from "child_process";
import { makeEstateBreakdownSpec } from "../src/visualTemplates/estateBreakdownSpec.js";

const values = {
  managedAssets: 72,
  realEstateValue: 28,
  otherAssets: 18
};

const spec = makeEstateBreakdownSpec(values);

fs.mkdirSync("public/generated-visuals", { recursive: true });
fs.writeFileSync(
  "public/generated-visuals/estate-breakdown.vl.json",
  JSON.stringify(spec, null, 2)
);

execFileSync(
  "npx",
  ["vl2svg", "public/generated-visuals/estate-breakdown.vl.json"],
  {
    stdio: [
      "ignore",
      fs.openSync("public/generated-visuals/estate-breakdown.svg", "w"),
      "inherit"
    ]
  }
);

console.log("Generated public/generated-visuals/estate-breakdown.svg");

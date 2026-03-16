const fs = require("fs");

const branch = process.env.BRANCH_OFFICE
  .normalize("NFD")
  .replace(/[\u0300-\u036f]/g,"")
  .replace(/\s/g,"_");

const raw = JSON.parse(
  fs.readFileSync(`report-data-${branch}.json`)
);

const clean = {};

for (const visual in raw.visuals) {

  const rows = raw.visuals[visual];

  if (!Array.isArray(rows)) continue;

  const parsed = rows.map(r => {

    const parts = r.split(/\s{2,}/);

    return {
      label: parts[0],
      value: parts[1] || ""
    };

  });

  clean[visual] = parsed;

}

fs.writeFileSync(
  `report-clean-${branch}.json`,
  JSON.stringify(clean,null,2)
);

console.log("Parsed visuals:",Object.keys(clean).length);

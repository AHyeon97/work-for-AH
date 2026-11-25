const { buildWorksheet, buildKey } = require("./render.js");
const DATA = require(process.argv[2] || "./content_a.js");
(async () => {
  for (const d of DATA) {
    await buildWorksheet(d, `늘봄앵커_우리충남탐험놀이_${d.file}_활동지.pptx`);
    await buildKey(d, `늘봄앵커_우리충남탐험놀이_${d.file}_활동지_해설지.pptx`);
    console.log("ok:", d.file);
  }
})();

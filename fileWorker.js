importScripts('https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js');

const stageOrder = ["INIT","AST","FLA","IOT","FCT","FC2","IST","FPF","NVL"];

self.onmessage = function(e) {

  const {buffers} = e.data;
  let results = [];

  buffers.forEach(buffer => {

    const workbook = XLSX.read(buffer, {type:'array'});
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet);

    let dur = {};
    let cnt = {};

    json.forEach(r=>{
      const s = r.Stage;
      const st = new Date(r.StartTime);
      const et = new Date(r.EndTime);

      if (s && !isNaN(st) && !isNaN(et)){
        const d = (et - st)/1000;

        if (!dur[s]) { dur[s]=0; cnt[s]=0; }

        dur[s]+=d;
        cnt[s]++;
      }
    });

    results.push({dur,cnt});
  });

  // ✅ send back result
  self.postMessage({results});
};
``
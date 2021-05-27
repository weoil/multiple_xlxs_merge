import xlsx from "node-xlsx";
import fs from "fs";
import path from "path";
import inquirer from "inquirer";
const yyjlData: any = {};
const kData: any = {};
async function readData(filePath: string, department: string) {
  const workSheetsFromFile = xlsx.parse(filePath);
  const ks: any = {};
  workSheetsFromFile.forEach((sheet, index) => {
    // K
    if (index == 0) {
      sheet.data.forEach((row: any[], i) => {
        let number = row[0];
        if (typeof number !== "number") {
          return;
        }
        if (!ks[row[3]]) {
          ks[row[3]] = {};
        }
        ks[row[3]]["k"] = {
          类型: row[2],
          投资圈名称: row[3],
          投资圈编号: row[4],
          "开拓人（BD）": row[5],
          圈主: row[6],
          专属投顾: row[7],
          是否达成合作意向: row[8],
          大搜股助手APP: row[9],
          是否建立投资圈群: row[10],
          级别: 1,
          是否合格: "不合格",
          所属业务部: "",
        };
      });
      return;
    }
    // C信息
    if (index == 1) {
      sheet.data.forEach((row: any[], i) => {
        if (row[13] !== "是") {
          return;
        }

        if (!ks[row[1]]) {
          ks[row[1]] = {};
        }
        if (!ks[row[1]]["c"]) {
          ks[row[1]]["c"] = [];
        }
        ks[row[1]]["c"].push({
          客户编号: row[2],
          客户微信: row[3],
          客户姓名: row[4],
          所在群: row[5],
          是否下载app: row[6],
          是否建立V群: row[7],
        });
        // console.log(ks[row[1]]["c"]);
      });
      return;
    }
    // C成交
    if (index == 2) {
      sheet.data.forEach((row: any[], i) => {
        if (row[4].length < 3) {
          return;
        }
        if (!ks[row[4]]) {
          ks[row[4]] = {};
        }
        if (!ks[row[4]]["m"]) {
          ks[row[4]]["m"] = [];
        }
        ks[row[4]]["m"].push({
          客户姓名: row[0],
          成交项目: row[1],
          成交金额: row[2],
          成交时间: row[3],
          所属业务部: row[5],
          是否清算: row[6],
        });
      });
      return;
    }
  });
  const xlsxData: any[] = [];
  // const template = xlsx.parse(`${__dirname}/template.xlsx`);
  const template = [
    ["投资圈基本信息", "", "", "", "", ""],
    ["类型", "", "投资圈名称", "", "投资圈编号", ""],
    ["开拓人（BD）", "", "圈主", "", "专属投顾", ""],
    ["是否达成合作意向", "", "大搜股助手APP", "", "是否建立投资圈群", ""],
    ["级别", "1级", "是否合格", "否", "所属业务部", ""],
    ["投资圈合格客户信息", "", "", "", "", ""],
    [
      "客户编号",
      "客户微信",
      "客户姓名",
      "所在群",
      "是否下载app",
      "是否建立V群",
    ],
  ];
  Object.keys(ks).forEach((kName) => {
    const obj = ks[kName];
    const data = JSON.parse(JSON.stringify(template));
    let rowCount = 0;
    if (!obj["k"]) {
      return;
    }
    // 填充K信息
    data[1][1] = obj["k"]["类型"];
    data[1][3] = obj["k"]["投资圈名称"];
    data[1][5] = obj["k"]["投资圈编号"];
    data[2][1] = obj["k"]["开拓人（BD）"];
    data[2][3] = obj["k"]["圈主"];
    data[2][5] = obj["k"]["专属投顾"];
    data[3][1] = obj["k"]["是否达成合作意向"];
    data[3][3] = obj["k"]["大搜股助手APP"];
    data[3][5] = obj["k"]["是否建立投资圈群"];
    data[4][1] = kData[kName]["级别"] || obj["k"]["级别"];
    data[4][3] = kData[kName]["是否合格"] || obj["k"]["是否合格"];
    data[4][5] = department;

    rowCount = 7;
    // 填充C信息
    if (!obj["c"]) {
      obj["c"] = [];
    }
    let cList: any[] = obj["c"];
    let cLen = cList.length;
    cLen = Math.max(cLen, 21);
    cList = cList.concat(getEmptyArray(21)).slice(0, cLen);

    cList.forEach((citem) => {
      if (!data[rowCount]) {
        data[rowCount] = [];
      }
      data[rowCount][0] = citem["客户编号"];
      data[rowCount][1] = citem["客户微信"];
      data[rowCount][2] = citem["客户姓名"];
      data[rowCount][3] = citem["所在群"];
      data[rowCount][4] = citem["是否下载app"];
      data[rowCount][5] = citem["是否建立V群"];
      rowCount += 1;
    });

    let cjxmRowIndex = rowCount++;
    data[cjxmRowIndex] = ["成交项目", "", "", "", "", ""];
    data[rowCount++] = [
      "客户姓名",
      "成交项目",
      "成交金额",
      "成交时间",
      "所属业务部",
      "是否清算",
    ];
    if (!obj["m"]) {
      obj["m"] = [];
    }
    let mList: any[] = obj["m"];
    let mLen = mList.length;
    mLen = Math.max(mLen, 26);
    mList = mList.concat(getEmptyArray(26)).slice(0, mLen);

    mList.forEach((mitem, cindex) => {
      if (!data[rowCount]) {
        data[rowCount] = [];
      }
      data[rowCount][0] = mitem["客户姓名"];
      data[rowCount][1] = mitem["成交项目"];
      data[rowCount][2] = mitem["成交金额"];
      data[rowCount][3] = mitem["成交时间"];
      data[rowCount][4] = mitem["所属业务部"];
      data[rowCount][5] = mitem["是否清算"];
      rowCount += 1;
    });
    let yyjlIndex = rowCount++;
    data[yyjlIndex] = ["运营记录", "", "", "", "", ""];
    data[rowCount++] = ["服务日期", "服务内容", "", "", "", ""];
    let yList: any[] = yyjlData[kName];
    if (!yList) {
      yList = [];
    }
    let yLen = yList.length;
    yLen = Math.max(yLen, 26);
    yList = yList.concat(getEmptyArray(26)).slice(0, yLen);

    data.push(...yList);
    const ops = {
      "!cols": [
        { wch: 15 },
        { wch: 15 },
        { wch: 15 },
        { wch: 15 },
        { wch: 15 },
        { wch: 15 },
      ],
      "!merges": [
        {
          s: {
            c: 0,
            r: 0,
          },
          e: {
            c: 5,
            r: 0,
          },
        },
        {
          s: {
            c: 0,
            r: 5,
          },
          e: {
            c: 5,
            r: 5,
          },
        },
        {
          s: {
            c: 0,
            r: cjxmRowIndex,
          },
          e: {
            c: 5,
            r: cjxmRowIndex,
          },
        },
        {
          s: {
            c: 0,
            r: yyjlIndex,
          },
          e: {
            c: 5,
            r: yyjlIndex,
          },
        },
        {
          s: {
            c: 1,
            r: yyjlIndex + 1,
          },
          e: {
            c: 5,
            r: yyjlIndex + 1,
          },
        },
      ],
    };
    let _yyContentStartIndex = yyjlIndex + 1;
    ops["!merges"].push(
      ...yList.map(() => {
        _yyContentStartIndex++;
        return {
          s: {
            c: 1,
            r: _yyContentStartIndex,
          },
          e: {
            c: 5,
            r: _yyContentStartIndex,
          },
        };
      })
    );
    xlsxData.push({
      name: kName,
      data: data,
      options: ops,
    });
  });
  const now = new Date();

  const outPath = path.resolve(
    path.dirname(filePath),
    `${department}-${now.getMonth()}${now.getDate()}${now.getHours()}${now.getMinutes()}.xlsx`
  );
  fs.writeFileSync(outPath, xlsx.build(xlsxData) as any);
  console.log("文件保存在:", outPath);
}

function getEmptyArray(len: number) {
  return new Array(len).fill(["", "", "", "", "", ""]);
}
function readYYJLData(filePath: string) {
  try {
    const workSheetsFromFile = xlsx.parse(filePath);
    workSheetsFromFile.forEach((i) => {
      kData[i.name] = {};
      const _old: any[] = [];
      let started = false;
      i.data.forEach((row, index) => {
        if (index === 4) {
          kData[i.name] = {
            级别: row[1],
            是否合格: row[3],
          };
        }
        if (started) {
          _old.push(row);
          return;
        }
        if (!started && row[0] === "服务日期" && row[1] === "服务内容") {
          started = true;
        }
      });
      yyjlData[i.name] = _old;
    });
  } catch (e) {}
}

async function main() {
  const sourceResult = await inquirer.prompt({
    type: "input",
    message: "请输入文件地址",
    name: "data",
  });
  console.log(sourceResult.data);
  let sourceUrl = sourceResult.data as string;
  sourceUrl = sourceUrl.replace(/\'/g, "").replace(/\"/g, "").trim();

  const oldFileResult = await inquirer.prompt({
    type: "input",
    message: "请输入上次生成的文件(无则留空)",
    name: "data",
  });
  let oldFilePath = oldFileResult.data as string;
  oldFilePath = oldFilePath.replace(/\'/g, "").replace(/\"/g, "").trim();
  const departmentResult = await inquirer.prompt({
    type: "input",
    message: "请输入该文件所属的业务部门",
    name: "data",
  });

  if (oldFilePath !== "") {
    readYYJLData(oldFilePath);
  }
  readData(sourceUrl, departmentResult.data);
}
main();

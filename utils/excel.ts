import * as XLSX from 'xlsx';

interface IExportExcelParameters {
  headers: any; // excel的头
  data: any; // 数据
  fileName?: string; // 导出文件名
  cols?: any[]; // 行间距
  merges?: any[]; // 合并行
}

/*
以下是使用示范

headers: any; // excel的头
const headers = [{
  title: '姓名',
  dataIndex: 'name',
  key: 'name',
}, {
  title: '年级',
  dataIndex: 'grade',
  key: 'grade',
}, {
  title: '部门',
  dataIndex: 'department',
  key: 'department',
  className: 'text-monospace',
}];

data: any; // 数据
let attendanceInfoList = [
  {
    name: '张三',
    grade: '2017级',
    department: '前端部门',
  },
  {
    name: '李四',
    grade: '2017级',
    department: '程序部门',
  },
];

fileName?: string; // 导出文件名
'工资单.xlsx'

cols?: any[]; // 行间距
[{wpx: 100},{wpx: 100}]

merges?: any[]; // 合并行
[
  {
    s: {
      c: 11,
      r: length,
    },
    e: {
      c: 15,
      r: length,
    },
  }
]
*/

// 导出excel
export const exportExcel = ({
                              headers,
                              data,
                              fileName = '表格.xlsx',
                              cols = [],
                              merges = [],
                            }: IExportExcelParameters) => {
  const _headers = headers
    .map((item: any, i: any) =>
      Object.assign(
        {},
        {
          key: item.dataIndex || item.key,
          title: item.title,
          position:
            Number(i + 1) > 26
              ? String.fromCharCode(Number(65 + (Math.floor(headers.length / 26) - 1))) +
              String.fromCharCode(65 + (i - 26)) +
              1
              : String.fromCharCode(65 + i) + 1,
        },
      ),
    )
    .reduce(
      (prev: any, next: any) => Object.assign({}, prev, { [next.position]: { key: next.key, v: next.title } }),
      {},
    );

  const _data = data
    .map((item: any, i: any) =>
      headers.map((key: any, j: any) =>
        Object.assign(
          {},
          {
            content: item[key.dataIndex || key.key],
            position:
              Number(j + 1) > 26
                ? String.fromCharCode(Number(65 + (Math.floor(headers.length / 26) - 1))) +
                String.fromCharCode(65 + (j - 26)) +
                (i + 2)
                : String.fromCharCode(65 + j) + (i + 2),
          },
        ),
      ),
    )
    .reduce(
      // 对刚才的结果进行降维处理（二维数组变成一维数组）
      (prev: any, next: any) => prev.concat(next),
    )
    .reduce(
      // 转换成 worksheet 需要的结构
      (prev: any, next: any) => Object.assign({}, prev, { [next.position]: { v: next.content } }),
      {},
    );

  // 合并 headers 和 data
  const output = Object.assign({}, _headers, _data);
  // 获取所有单元格的位置
  const outputPos = Object.keys(output);
  // 计算出范围 ,["A1",..., "H2"]
  const ref = `${outputPos[0]}:${outputPos[outputPos.length - 1]}`;

  // 构建 workbook 对象
  const wb = {
    SheetNames: ['mySheet'],
    Sheets: {
      mySheet: Object.assign({}, output, {
        '!ref': ref,
        '!cols': cols,
        '!merges': merges,
      }),
    },
  };

  // 导出 Excel
  XLSX.writeFile(wb, fileName);
};

/**
 * 将 excel 读取为 json
 */
export const readXlsxAsJson = async (file: File) =>
  new Promise<Record<string, string>[]>((resolve) => {
    const reader = new FileReader();
    reader.onload = (e: ProgressEvent<any>) => {
      const data = new Uint8Array(e.target?.result);
      const workbook = XLSX.read(data, { type: 'array' });

      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonArr = XLSX.utils.sheet_to_json(worksheet);

      resolve(jsonArr as any);
    };
    reader.readAsArrayBuffer(file);
  });

import { isObject } from "lodash";
import { Column } from "../interfaces";

export const exportToExcel = async (
  columns: Column[],
  fileName: string,
  exportExamples?: Record<string, string>[]
): Promise<void> => {
  const { utils: XLSXUtils, writeFile: XLSXWriteFile } = await import(
    /* webpackChunkName: "chunks/xlsx.015f" */
    "xlsx"
  );

  const headers = columns.map((col) => ({
    key: col.name,
    header: col.label || col.name,
  }));

  const data = exportExamples || [];
  const worksheet = XLSXUtils.json_to_sheet(data, {
    header: headers.map((h) => h.header),
  });

  const workbook = XLSXUtils.book_new();
  XLSXUtils.book_append_sheet(workbook, worksheet, "Template");
  XLSXWriteFile(workbook, `${fileName}.xlsx`);
};

// 新增导出表单数据的函数
export const exportFormData = async (
  columns: Column[],
  formData: Record<string, any>[],
  fileName: string
): Promise<void> => {
  const { utils: XLSXUtils, writeFile: XLSXWriteFile } = await import(
    /* webpackChunkName: "chunks/xlsx.015f" */
    "xlsx"
  );

  // 处理表头
  const headers = columns.map((col) => ({
    key: col.name,
    header: col.label || col.name,
  }));

  // 处理数据，确保数据格式正确
  const exportData = formData.map((row) => {
    const exportRow: Record<string, any> = {};
    headers.forEach((header) => {
      exportRow[header.header] = row[header.key];
    });
    return exportRow;
  });

  // 创建工作表和工作簿
  const worksheet = XLSXUtils.json_to_sheet(exportData);
  const workbook = XLSXUtils.book_new();
  XLSXUtils.book_append_sheet(workbook, worksheet, "Data");
  XLSXWriteFile(workbook, `${fileName}.xlsx`);
};

export const importFromExcel = async (
  file: File,
  columns: Column[],
  importFilter?: string
): Promise<Record<string, any>[]> => {
  // sha1 hash of "dynamic-form-item-v2" starts with "015f"
  const { utils: XLSXUtils, read: XLSXRead } = await import(
    /* webpackChunkName: "chunks/xlsx.015f" */
    "xlsx"
  );

  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      if (!e.target?.result) {
        reject(new Error("Failed to read file"));
        return;
      }

      const data = new Uint8Array(e.target.result as ArrayBuffer);
      const workbook = XLSXRead(data, { type: "array" });

      if (!workbook.SheetNames.length) {
        reject(new Error("No sheets found in workbook"));
        return;
      }

      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData: Record<string, string>[] = XLSXUtils.sheet_to_json(
        worksheet,
        {
          raw: false, // 返回格式化的字符串
        }
      );

      // 增强数据验证和转换逻辑
      let transformedData = jsonData.map((row: Record<string, string>) => {
        const transformedRow: Record<string, any> = {};

        columns.forEach((col) => {
          const header = col.label || col.name;
          const value = row[header];
          // 添加类型检查和转换
          transformedRow[col.name] = validateAndTransformValue(value, col);
        });

        return transformedRow;
      });

      // 如果提供了ImportFilter，则根据指定字段进行去重
      // ========== 新增：根据 importFilter 字段的 options 进行过滤 ==========
      if (importFilter && importFilter.length > 0) {
        // 构建字段配置映射
        const filterFieldConfigs = importFilter
          .split(",")
          .map((fieldName) => {
            const column = columns.find((col) => col.name === fieldName);
            if (!column) {
              return null;
            }

            // 获取 options 列表
            let options: any[] = [];
            if (column.type === "select" && column.props?.options) {
              options = Array.isArray(column.props.options)
                ? column.props.options
                : [];
            }

            return {
              fieldName,
              column,
              options,
            };
          })
          .filter((config) => config !== null);

        // 第一步：根据 options 过滤无效数据
        transformedData = transformedData.filter((row) => {
          // 检查所有 filter 字段
          for (const config of filterFieldConfigs) {
            const value = row[config.fieldName];

            // 如果值为空，保留该行（允许空值）
            if (value === undefined || value === null || value === "") {
              continue;
            }

            // 如果 options 为空数组，不过滤（允许所有值）
            if (config.options.length === 0) {
              continue;
            }

            // 检查 value 是否在 options 中
            const isValid = config.options.some((option) => {
              // 支持不同的 option 结构
              const optionValue =
                option.value !== undefined ? option.value : option;
              return String(optionValue) === String(value);
            });

            // 如果值不在 options 中，过滤掉该行
            if (!isValid) {
              return false;
            }
          }
          return true;
        });

        // 第二步：根据 importFilter 字段组合进行去重
        const uniqueData: Record<string, any>[] = [];
        const seen = new Set<string>();

        transformedData.forEach((row) => {
          // 构建用于判断重复的键
          const filterKey = importFilter
            .split(",")
            .map((field) => {
              const value = row[field];
              return value === undefined || value === null
                ? "__NULL__"
                : String(value);
            })
            .join("|");

          // 如果未见过这个键，则添加到结果中
          if (!seen.has(filterKey)) {
            seen.add(filterKey);
            uniqueData.push(row);
          }
        });

        transformedData = uniqueData;
      }

      resolve(transformedData);
    };

    reader.onerror = reject;

    reader.readAsArrayBuffer(file);
  });
};

export function validateAndTransformValue(
  value: string | undefined,
  column: Column
): any {
  // 如果值为空，则返回空值
  const trimValue = value?.trim();
  if (trimValue === undefined || trimValue === "") {
    return undefined;
  }

  // 根据类型，处理不同的值
  switch (column.type) {
    case "cascader":
      try {
        const parsedValue = JSON.parse(trimValue);
        if (Array.isArray(parsedValue)) {
          return parsedValue;
        }
      } catch {
        // eslint-disable-next-line no-empty
      }
      // 尝试按 / 分割
      return trimValue.split("/").map((v) => v.trim());

    case "inputNumber":
      // 兼容非数字(NaN)
      return Number(trimValue) || 0;

    case "select":
      if (column.props?.mode === "multiple") {
        try {
          const parsedValue = JSON.parse(trimValue);
          if (Array.isArray(parsedValue)) {
            return parsedValue;
          }
        } catch {
          // eslint-disable-next-line no-empty
        }
        // 如果换行符数量小于等于1，则按逗号分割
        const trySplitLines = trimValue.split("\n").map((v) => v.trim());
        if (trySplitLines.length <= 1) {
          return trimValue.split(",").map((v) => v.trim());
        }
        return trySplitLines;
      } else {
        return trimValue;
      }

    case "checkbox": {
      const trimmedValue = trimValue.toLowerCase();
      // 明确的 false 值
      if (
        trimmedValue === "false" ||
        trimmedValue === "0" ||
        trimmedValue === "no" ||
        trimmedValue === "n" ||
        trimmedValue === "否"
      ) {
        return false;
      }
      // 其他所有情况都是 true
      return true;
    }

    case "timeRangePicker": {
      try {
        const parsedValue = JSON.parse(trimValue);
        if (
          isObject(parsedValue) &&
          "startTime" in parsedValue &&
          "endTime" in parsedValue
        ) {
          return parsedValue;
        }
      } catch {
        // eslint-disable-next-line no-empty
      }
      // 尝试按 ~ 分割
      let trySplit = trimValue.split("~").map((v) => v.trim());
      if (trySplit.length <= 1) {
        trySplit = trimValue.split(",").map((v) => v.trim());
      }
      const [startTime, endTime] = trySplit;
      return { startTime, endTime };
    }

    case "textArea":
    case "input":
    case "inputPassword":
    case "autoComplete":
    default:
      return trimValue;
  }
}

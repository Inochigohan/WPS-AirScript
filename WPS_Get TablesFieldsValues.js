/**
 * 根据表名和字段名获取表ID、字段ID及字段所有行的值（AirScript实现）
 * 需在WPS脚本编辑器中运行
 * 参考官方文档：https://airsheet.wps.cn/docs/api/dbsheet/Field.html
 */

// 获取所有表信息
const sheets = Application.Sheet.GetSheets();
// console.log(sheets);

function getTableFieldInfo(tableName, fieldName) {
    let tableId = null;
    let fieldId = null;
    let fieldValues = [];

    // 遍历所有表，查找目标表
    for (let i = 0; i < sheets.length; i++) {
        const sheet = sheets[i];
        if (sheet.name === tableName) {
            tableId = sheet.id;
            // 遍历表的字段，查找目标字段
            for (let j = 0; j < sheet.fields.length; j++) {
                const field = sheet.fields[j];
                if (field.name === fieldName) {
                    fieldId = field.id;
                    break;
                }
            }
            break;
        }
    }

    // 如果未找到表或字段，直接返回
    if (tableId === null || fieldId === null) {
        return {
            tableId: tableId,
            fieldId: fieldId,
            fieldValues: fieldValues
        };
    }

    // 获取表的所有记录
    let offset = null;
    let allRecords = [];
    do {
        const records = Application.Record.GetRecords({
            SheetId: tableId,
            Offset: offset
        });
        allRecords = allRecords.concat(records.records);
        offset = records.offset;
    } while (offset);
    // console.log(allRecords);

    // 提取目标字段每一行的值
    for (let k = 0; k < allRecords.length; k++) {
        const record = allRecords[k];
        const fieldValue = record.fields[fieldName];
        fieldValues.push(fieldValue);
    }

    return {
        tableId: tableId,
        fieldId: fieldId,
        fieldValues: fieldValues
    };
}

// 示例输入
const tableName = "信息同步";
const fieldName = "问题关键字";
const result = getTableFieldInfo(tableName, fieldName);
console.log(`表名: ${tableName}`);
console.log(`表 ID: ${result.tableId}`);
console.log(`字段名: ${fieldName}`);
console.log(`字段 ID: ${result.fieldId}`);
console.log(`该字段每一行的值:`, result.fieldValues);
function createTableAndFields() {
    const TARGET_SHEET_NAME = "小林动态报告";
    let CURRENT_SHEET_ID = null; 

    // ================== 动态创建/获取数据表 ==================
    const sheets = Application.Sheet.GetSheets();
    let targetSheet = null;

    // 查找目标表
    for (const sheet of sheets) {
        if (sheet.name === TARGET_SHEET_NAME) {
            targetSheet = sheet;
            break;
        }
    }

    // 表存在性处理
    if (targetSheet) {
        CURRENT_SHEET_ID = targetSheet.id;
        console.log(`[INFO] 表已存在，ID：${CURRENT_SHEET_ID}`);
    } else {
        // 创建新表
        const newSheet = Application.Sheet.CreateSheet({
            Name: TARGET_SHEET_NAME,
            Views: [{ name: '表格视图', type: 'Grid' }],
            Fields: [
                { 
                    name: '工单key',
                    type: 'MultiLineText'
                }
            ]
        });
        CURRENT_SHEET_ID = newSheet.id;
        console.log(`[SUCCESS] 数据表创建成功，ID：${CURRENT_SHEET_ID}`);
    }

    // ================== 创建字段结构 ==================
    const fieldDefinitions = [
        // 多选字段（带预定义选项）
        {
            name: '工单类型',
            type: 'SingleSelect',
            items: [
                { value: '问题记录' },
                { value: '点名批评' },
                { value: '奖赏表扬' },
                { value: '待办事项' },
                { value: '日常行为' },
                { value: '习惯复盘' }
            ]
        },
        { name: '工单概要', type: 'MultiLineText' },
        { name: '客户名称', type: 'MultiLineText' },
        { name: '创建日期', type: 'Date' },
        {
            name: '优先级',
            type: 'SingleSelect',
            items: [
                { value: '最高' },
                { value: '高' },
                { value: '普通' },
                { value: '低' },
                { value: '较低' }
            ]
        },
        { name: '报告人', type: 'MultiLineText' },
        { name: '经办人', type: 'MultiLineText' },
        {
            name: '状态',
            type: 'SingleSelect',
            items: [
                { value: '开放' },
                { value: '处理中' },
                { value: '已解决' },
                { value: '重新打开' }
            ]
        },
        { name: '解决结果', type: 'MultiLineText' },
        { name: '已解决时间', type: 'MultiLineText' }       
    ];
    // 新增字段存在性检查
    if (fieldDefinitions.length > 0) {
        // 执行创建操作
        const creationResult = Application.Field.CreateFields({
                SheetId: CURRENT_SHEET_ID,
                Fields: fieldDefinitions
            });
            console.log('[SUCCESS] 新增字段创建结果：', creationResult);
    } else {
        console.log('[WARN] 未定义需要创建的字段');
    }
}

// 执行函数
createTableAndFields();

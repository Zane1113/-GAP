/**

* @OnlyCurrentDoc

*/



// 在电子表格中创建一个自定义菜单

function onOpen() {

SpreadsheetApp.getUi()

.createMenu('Giao')

.addItem('Update storage', 'extractEmailContent')

.addToUi();

}



/**

* 主函数：抓取Gmail邮件内容并记录到表格

*/

function extractEmailContent() {

const sheetName = 'records';

const ss = SpreadsheetApp.getActiveSpreadsheet();

let sheet = ss.getSheetByName(sheetName);



if (!sheet) {

sheet = ss.insertSheet(sheetName);

}



// 读取已存在记录，按“Merchant + Invoice Amount + Email Titile”构建去重键

const existingKeys = new Set();

const lastRow = sheet.getLastRow();

if (lastRow >= 2) {

const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

const merchantIdx = headers.indexOf('Merchant');

const amountIdx = headers.indexOf('Invoice Amount');

const titleIdx = headers.indexOf('Email Titile');

if (merchantIdx !== -1 && amountIdx !== -1 && titleIdx !== -1) {

const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

rows.forEach(r => {

const merchant = String(r[merchantIdx] || '').trim();

const amountRaw = String(r[amountIdx] || '').trim();

const amountKey = normalizeAmountForKey_(amountRaw);

const title = String(r[titleIdx] || '').trim();

if (merchant || amountKey || title) {

const key = `${merchant}|${amountKey}|${title}`;

existingKeys.add(key);

}

});

}

}



// 使用标签对象来获取线程（避免 label:"..." 在脚本中因空格/层级导致搜索不到）

const labelPath = '0-system/[Invoice Request]'; // 按你的实际标签路径

const label = GmailApp.getUserLabelByName(labelPath);

if (!label) {

SpreadsheetApp.getUi().alert('未找到该标签：' + labelPath + '。请检查：大小写、层级分隔符“/”、是否为用户自定义标签。');

Logger.log('Label not found: ' + labelPath);

return;

}



// 分页取该标签下所有线程

let threads = [];

let start = 0;

const pageSize = 100;

while (true) {

const batch = label.getThreads(start, pageSize);

if (batch.length === 0) break;

threads = threads.concat(batch);

start += batch.length;

if (batch.length < pageSize) break;

}

Logger.log('使用标签路径: ' + labelPath);

Logger.log('找到邮件线程数量: ' + threads.length);



const allData = [];

let processedCount = 0;

let skippedCount = 0;



for (const thread of threads) {

const messages = thread.getMessages();

if (messages.length === 0) continue;



// 线程级：汇总 PDF 附件名（去重）

const threadPdfNames = gatherThreadPdfNames_(thread);



// 逐消息处理

for (const message of messages) {

const sender = message.getFrom();

const subject = message.getSubject();



// 新的判定规则：发件人满足 或 该邮件带有PDF，则处理；否则跳过

const hasPdf = messageHasPdf_(message);

const senderAllowed = sender.includes('it@klook.com') || sender.includes('supply.tech@klook.com');

if (!senderAllowed && !hasPdf) {

Logger.log('跳过邮件（发件人不匹配且无PDF）：' + sender + ' | 主题: ' + subject);

skippedCount++;

continue;

}



try {

const emailDate = message.getDate();

const bodyText = message.getPlainBody();

const bodyHtml = message.getBody();



// 解析所有字段（文本冒号 + HTML表格）

const extractedAll = extractFieldsFromEmail(bodyText, bodyHtml);



// 构建“选定字段”记录（含 Email Titile、Request Date、Invoice Date 等）

const record = buildSelectedRecord_({

rawFields: extractedAll,

emailSubject: subject,

sender,

emailDate,

bodyText,

threadPdfNames,

hasPdf,

bodyHtml, // 将HTML内容也传递给构建记录函数

});



// 去重键：Merchant + Invoice Amount + Email Titile（金额使用归一化）

const merchant = String(record['Merchant'] || '').trim();

const amountKey = normalizeAmountForKey_(record['Invoice Amount'] || '');

const title = String(record['Email Titile'] || '').trim();

const dedupKey = `${merchant}|${amountKey}|${title}`;

if (existingKeys.has(dedupKey)) {

Logger.log('跳过重复（Merchant+Invoice Amount+Email Titile）: ' + dedupKey);

skippedCount++;

continue;

}

existingKeys.add(dedupKey);



allData.push(record);

processedCount++;

Logger.log('处理邮件: ' + subject + ' - 字段数: ' + Object.keys(record).length);

} catch (error) {

Logger.log('处理邮件时出错: ' + error.toString());

}

}

}



Logger.log('处理统计 - 处理成功: ' + processedCount + ', 跳过: ' + skippedCount);



if (allData.length > 0) {

writeDataToSheet(sheet, allData);

SpreadsheetApp.getUi().alert('处理完成！成功提取 ' + allData.length + ' 封邮件的内容。');

} else {

SpreadsheetApp.getUi().alert('没有找到符合条件的邮件。');

}

}



/**

* 从邮件内容中提取字段和内容

*/

function extractFieldsFromEmail(bodyText, bodyHtml) {

const fields = {};



// 1. 从纯文本中提取带冒号的字段

const textFields = extractFieldsFromText(bodyText);

Object.assign(fields, textFields);



// 2. 从HTML表格中提取字段

const tableFields = extractFieldsFromTable(bodyHtml);

Object.assign(fields, tableFields);



return fields;

}



/**

* 从纯文本中提取字段（格式：字段名: 内容）

*/

function extractFieldsFromText(bodyText) {

const fields = {};

const lines = bodyText.split('\n');



for (const line of lines) {

// 查找包含冒号的行

const colonIndex = line.indexOf(':');

if (colonIndex > 0 && colonIndex < line.length - 1) {

const fieldName = line.substring(0, colonIndex).trim();

const fieldValue = line.substring(colonIndex + 1).trim();



// 过滤掉一些常见的非字段行（如时间、邮箱地址等）

if (fieldName && fieldValue && !isSystemField(fieldName)) {

fields[fieldName] = fieldValue;

}

}

}



return fields;

}



/**

* 从HTML表格中提取字段

*/

function extractFieldsFromTable(bodyHtml) {

const fields = {};



try {

const tableRowRegex = /<tr[^>]*>(.*?)<\/tr>/gi;

let tableMatch;



while ((tableMatch = tableRowRegex.exec(bodyHtml)) !== null) {

const rowHtml = tableMatch[1];


const cellRegex = /<td[^>]*>(.*?)<\/td>/gi;

const cells = [];

let cellMatch;



while ((cellMatch = cellRegex.exec(rowHtml)) !== null) {

const cellText = cellMatch[1].replace(/<[^>]+>/g, '').trim();

if (cellText) {

cells.push(cellText);

}

}



if (cells.length >= 2) {

const fieldName = cells[0];

const fieldValue = cells[1];



if (fieldName && fieldValue && !isSystemField(fieldName)) {

fields[fieldName] = fieldValue;

}

}

}

} catch (error) {

Logger.log('解析HTML表格时出错: ' + error.toString());

}



return fields;

}



/**

* 判断是否为系统字段（需要过滤的字段）

*/

function isSystemField(fieldName) {

const systemFields = [

'from', 'to', 'cc', 'bcc', 'date', 'time', 'subject',

'发件人', '收件人', '抄送', '日期', '时间', '主题',

'gmail', 'email', 'message-id'

];



return systemFields.some(sysField =>

fieldName.toLowerCase().includes(sysField.toLowerCase())

);

}



/**

* 将提取的数据写入表格

*/

function writeDataToSheet(sheet, allData) {

// 收集所有唯一的字段名

const allFieldNames = new Set();

allData.forEach(data => {

Object.keys(data).forEach(key => allFieldNames.add(key));

});



const fieldNames = Array.from(allFieldNames).sort();



// 如果表格为空，添加表头

if (sheet.getLastRow() === 0) {

sheet.appendRow(fieldNames);

sheet.getRange(1, 1, 1, fieldNames.length).setFontWeight('bold');

} else {

// 如果表格已有数据，检查是否需要添加新列

const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

const newFields = fieldNames.filter(field => !existingHeaders.includes(field));


if (newFields.length > 0) {

// 添加新的列标题

const startCol = sheet.getLastColumn() + 1;

sheet.getRange(1, startCol, 1, newFields.length).setValues([newFields]);

sheet.getRange(1, startCol, 1, newFields.length).setFontWeight('bold');

}

}



// 获取当前的表头

const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];



// 准备数据行

const dataRows = allData.map(data => {

return currentHeaders.map(header => data[header] || '');

});



// 写入数据

if (dataRows.length > 0) {

const startRow = sheet.getLastRow() + 1;

sheet.getRange(startRow, 1, dataRows.length, currentHeaders.length).setValues(dataRows);

}



// 写入完成后，按 Request Date 顺序排序（表头不参与）

const headers = currentHeaders;

const reqIdx = headers.indexOf('Request Date');

if (reqIdx !== -1) {

const totalRows = sheet.getLastRow();

const totalCols = sheet.getLastColumn();

if (totalRows >= 2) {

sheet.getRange(2, 1, totalRows - 1, totalCols).sort([{ column: reqIdx + 1, ascending: true }]);

}

}



// 自动调整列宽

for (let i = 1; i <= sheet.getLastColumn(); i++) {

sheet.autoResizeColumn(i);

}

}



/**

* 构建选定字段与辅助函数

*/

function buildSelectedRecord_(ctx) {

const {

rawFields,

emailSubject,

sender,

emailDate,

bodyText,

threadPdfNames,

hasPdf,

bodyHtml, // 新增，用来在更底层的HTML源码中搜索

} = ctx;



// **针对特定字段的重构逻辑**

const commissionPeriod = extractCommissionPeriodFromHtml_(bodyHtml);

const invoiceRecipients = extractContactEmailFromHtml_(bodyHtml);

const receivedType = extractReceivedTypeFromHtml_(bodyHtml);



// 其他字段保持原有逻辑

const merchantStr = pickField_(rawFields, [

'Merchant ID and NAME that the invoice for',

'Merchant',

'Merchant ID and NAME'

]) || '';



const invoiceCurrency = pickField_(rawFields, [

'Currency',

'Invoice Currency'

]);



const invoiceAmount = pickField_(rawFields, [

'Invoice Amount',

'Amount'

]);



const paymentTerms = pickField_(rawFields, [

'Payment Terms (Days)',

'Payment terms'

]);



const companyName = pickField_(rawFields, [

'Full Company Name',

'Company Name'

]);



const address = pickField_(rawFields, [

'Address',

'Company Address'

]);



const receivableStart = pickField_(rawFields, [

'Receivable Timeframe (Starting Date)',

'Receivable Timeframe (Start Date)'

]);



const applicant = pickField_(rawFields, [

'Submitter',

'Applicant'

]);



// Business scenario description 支持多行

const otherCommentsMulti = extractMultiLineField_(bodyText, 'Business scenario description');

const otherComments = otherCommentsMulti || pickField_(rawFields, [

'Business scenario description',

'Other Comments (Business scenario description)'

]);



const invoicingParty = pickField_(rawFields, [

'Invoicing party'

]);



// 衍生字段：Merchant ID（从 Merchant 文本中提取数字）

const merchantId = extractDigits_(merchantStr);



// Request Date = 邮件发送日（yyyy-MM-dd）

const requestDate = formatDateYMD_(emailDate);



// Due Date(if have) = Request Date + INT(Payment terms)

let dueDateIfHave = '';

const termsDays = paymentTerms ? parseInt(extractDigits_(paymentTerms) || '0', 10) : 0;

if (termsDays > 0 && requestDate) {

const reqDateObj = emailDate instanceof Date ? emailDate : safeParseDate_(requestDate);

if (reqDateObj) {

dueDateIfHave = formatDateYMD_(addDays_(reqDateObj, termsDays));

}

}



// Invoice Date = 本封邮件含PDF时，取本封邮件的发送日期；否则留空

const invoiceDate = hasPdf ? formatDateYMD_(emailDate) : '';



// “邮件名” -> Email Titile（按你的命名）

const emailTitleField = emailSubject || '';



// “邮件线程中pdf附件附件名”

const threadPdfNamesField = threadPdfNames || '';



// 对展示的金额做千分位格式化（去重键会用归一化后的数值字符串，不受影响）

const invoiceAmountDisplay = formatAmount_(invoiceAmount || '');



// 汇总结果（按你的字段名）

const record = {

// 以表格区分的字段

'Merchant': merchantStr || '',

'Received Type': receivedType || '',

'Invoice Currency': invoiceCurrency || '',

'Invoice Amount': invoiceAmountDisplay || '',

'Payment terms': paymentTerms || '',

'Due Date(if have)': dueDateIfHave || '',

'Merchant Company Name': companyName || '',

'Merchant Address': address || '',

'Invoice Recipients(email)': invoiceRecipients || '',

'Commission period': commissionPeriod || '',



// 以“：”区分的字段

'Applicant': applicant || '',

'Other Comments (Business scenario description)': otherComments || '',

'Invoicing party': invoicingParty || '',



// 其他字段

'Merchant ID': merchantId || '',

'Email Titile': emailTitleField,

'PDF附件名': threadPdfNamesField,

'Invoice Date': invoiceDate,

'Request Date': requestDate

};



return record;

}



/**

* 线程级：汇总 PDF 附件名（去重）

*/

function gatherThreadPdfNames_(thread) {

const names = new Set();

try {

const messages = thread.getMessages();

messages.forEach(msg => {

let atts = [];

try {

atts = msg.getAttachments({ includeInlineImages: false, includeAttachments: true }) || [];

} catch (e) {

Logger.log('getAttachments error in gatherThreadPdfNames_: ' + e);

}

atts.forEach(a => {

if (!a) return;

let name = '';

let ctype = '';

try {

if (a.getName) name = a.getName() || '';

} catch (e) {

// ignore

}

// Prefer name-based check first

if (name && name.toLowerCase().endsWith('.pdf')) {

names.add(name);

return;

}

try {

if (a.getContentType) ctype = a.getContentType() || '';

} catch (e) {

// ignore contentType failures

}

if (ctype && String(ctype).toLowerCase().indexOf('pdf') !== -1) {

names.add(name || 'unnamed.pdf');

}

});

});

} catch (e) {

Logger.log('gatherThreadPdfNames_ error: ' + e);

}

return Array.from(names).join(', ');

}



/**

* 判断是否包含 PDF 附件

*/

function messageHasPdf_(message) {

try {

let atts = [];

try {

atts = message.getAttachments({ includeInlineImages: false, includeAttachments: true }) || [];

} catch (e) {

Logger.log('getAttachments error in messageHasPdf_: ' + e);

atts = message.getAttachments() || []; // fallback

}



for (let i = 0; i < atts.length; i++) {

const a = atts[i];

if (!a) continue;

let name = '';

let ctype = '';

try {

if (a.getName) name = a.getName() || '';

} catch (e) {

// ignore

}

if (name && name.toLowerCase().endsWith('.pdf')) return true;



try {

if (a.getContentType) ctype = a.getContentType() || '';

} catch (e) {

// ignore contentType failures

}

if (ctype && String(ctype).toLowerCase().indexOf('pdf') !== -1) return true;

}

} catch (e) {

Logger.log('messageHasPdf_ error: ' + e);

}

return false;

}



/**

* 提取多行字段

*/

function extractMultiLineField_(bodyText, label) {

if (!bodyText || !label) return '';



const knownLabels = [

'Submitter',

'Business scenario description',

'Invoicing party',

'Invoice Date',

'Request Date',

'Payment Terms (Days)',

'Full Company Name',

'Address',

'Contact Email',

'Receivable Timeframe (Starting Date)',

'By Cash/Credit/Deduct from payment',

'Currency',

'Invoice Amount'

];



const pattern = new RegExp(`${label}\\s*:\\s*([\\s\\S]*?)$`, 'i');

const tailMatch = bodyText.match(pattern);

if (!tailMatch) return '';



let content = tailMatch[1];



// 截断到下一个已知标签或两个连续换行

const nextLabelRegex = new RegExp(`\\n\\s*(?:${knownLabels.map(l => l.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\$&')).join('|')})\\s*:`, 'i');

const nlIdx = content.search(nextLabelRegex);

if (nlIdx >= 0) {

content = content.substring(0, nlIdx);

} else {

const dblNewlineIdx = content.indexOf('\n\n');

if (dblNewlineIdx >= 0) {

content = content.substring(0, dblNewlineIdx);

}

}



return content.trim();

}



/**

* 从字符串中提取数字

*/

function extractDigits_(str) {

if (!str) return '';

const m = String(str).match(/\d+/g);

return m ? m.join('') : '';

}



/**

* 格式化日期为 YYYY-MM-DD

*/

function formatDateYMD_(dateObj) {

try {

return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');

} catch (e) {

return '';

}

}



/**

* 安全地解析日期字符串

*/

function safeParseDate_(val) {

if (!val) return null;

if (val instanceof Date && !isNaN(val)) return val;



const s = String(val).trim();



// 标准解析

const d1 = new Date(s);

if (!isNaN(d1)) return d1;



// dd-MMM-yyyy（如 01-Jul-2025）

let m = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);

if (m) {

const day = parseInt(m[1], 10);

const monMap = {Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11};

const mon = monMap[m[2]];

const year = parseInt(m[3], 10);

if (mon !== undefined) {

const d = new Date(year, mon, day);

if (!isNaN(d)) return d;

}

}



// dd/MM/yyyy

m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);

if (m) {

const day = parseInt(m[1], 10);

const mon = parseInt(m[2], 10) - 1;

const year = parseInt(m[3], 10);

const d = new Date(year, mon, day);

if (!isNaN(d)) return d;

}



// yyyy-MM-dd

m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);

if (m) {

const year = parseInt(m[1], 10);

const mon = parseInt(m[2], 10) - 1;

const day = parseInt(m[3], 10);

const d = new Date(year, mon, day);

if (!isNaN(d)) return d;

}



return null;

}



/**

* 日期加天数

*/

function addDays_(dateObj, days) {

const d = new Date(dateObj.getTime());

d.setDate(d.getDate() + days);

return d;

}



/**

* 提取佣金月份

*/

function extractCommissionPeriod_(s) {

const d = safeParseDate_(s);

if (!d) return '';

const y = d.getFullYear();

const m = (d.getMonth() + 1).toString().padStart(2, '0');

return `${y}-${m}`;

}



/**

* 在已解析字段中按候选字段名列表依次查找，返回第一个非空值

*/

function pickField_(fields, candidateNames) {

if (!fields || !candidateNames || !candidateNames.length) return '';



const normalizedFields = {};

Object.keys(fields).forEach(k => {

normalizedFields[normalizeFieldName_(k)] = fields[k];

});



for (const cand of candidateNames) {

const normCand = normalizeFieldName_(cand);

if (normCand in normalizedFields) {

const v = normalizedFields[normCand];

if (v != null && String(v).trim() !== '') {

return String(v).trim();

}

}

}



return '';

}



/**

* 去重键用的金额归一化：去掉逗号、空格、货币符号，仅保留数字与小数点

*/

function normalizeAmountForKey_(s) {

return String(s || '')

.replace(/[,\s]/g, '')

.replace(/[^\d.]/g, ''); // 去掉除数字和点之外的字符

}



/**

* 金额展示格式化：给整数部分加千分位，保留原有小数位

*/

function formatAmount_(s) {

const raw = String(s || '').trim();

if (!raw) return '';

const normalized = normalizeAmountForKey_(raw);

const m = normalized.match(/^(\d+)(\.\d+)?$/);

if (!m) return raw; // 无法识别为数值，原样返回

let intPart = m[1];

const decPart = m[2] || '';

intPart = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',');

return intPart + decPart;

}



/**

* 字段名归一化

*/

function normalizeFieldName_(s) {

if (s == null) return '';

var out = String(s).trim();

out = out.toLowerCase();



// 移除所有空格和特殊字符，只保留字母和数字

out = out.replace(/[^a-z0-9]/g, '');



return out;

}



/**

* 清理 Received Type 字段

*/

function sanitizeReceivedType_(s) {

if (!s) return '';

return String(s).replace(/\s*[-–—|:]+$/g, '').trim();

}



// **专门为 Commission Period 字段重构的逻辑**

function extractCommissionPeriodFromHtml_(htmlBody) {

const monthMap = { jan: '01', feb: '02', mar: '03', apr: '04', may: '05', jun: '06', jul: '07', aug: '08', sep: '09', oct: '10', nov: '11', dec: '12' };


// 查找包含 "Receivable Timeframe (Starting Date)" 的单元格

const regex = /Receivable Timeframe \(Starting Date\)(?:[\s\S]*?)<td[^>]*>([\s\S]*?)<\/td>/i;

const match = htmlBody.match(regex);


if (match && match[1]) {

const dateStr = match[1].replace(/<[^>]+>/g, '').trim();

const parts = dateStr.match(/\d{1,2}-([a-z]{3})-(\d{4})/i);


if (parts && parts.length === 3) {

const year = parts[2];

const month = monthMap[parts[1].toLowerCase()];

if (month) {

return `${year}-${month}`;

}

}

}

return '';

}



// **专门为 Contact Email 字段重构的逻辑**

function extractContactEmailFromHtml_(htmlBody) {

// 查找包含 "Contact Email" 的单元格，并获取其下一个单元格的内容

const regex = /Contact Email(?:[\s\S]*?)<td[^>]*>([\s\S]*?)<\/td>/i;

const match = htmlBody.match(regex);


if (match && match[1]) {

// 移除所有HTML标签，包括<a>标签，并清理多余空格

const emailStr = match[1].replace(/<[^>]+>/g, '').trim();

const emailMatch = emailStr.match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/);

if (emailMatch) {

return emailMatch[0];

}

}

return '';

}



// **专门为 Received Type 字段重构的逻辑**

function extractReceivedTypeFromHtml_(htmlBody) {

// 查找包含 "By Cash/Credit/Deduct from payment" 的单元格，并获取其下一个单元格的内容

const regex = /By Cash\/Credit\/Deduct from payment(?:[\s\S]*?)<td[^>]*>([\s\S]*?)<\/td>/i;

const match = htmlBody.match(regex);


if (match && match[1]) {

// 移除所有HTML标签，并清理多余空格及末尾的标点符号

let typeStr = match[1].replace(/<[^>]+>/g, '').trim().replace(/[-–—|:.]*$/, '');



// 检查是否包含特定关键词

if (typeStr.toLowerCase().includes('credit')) {

return 'Credit';

}

if (typeStr.toLowerCase().includes('cash')) {

return 'Cash';

}

if (typeStr.toLowerCase().includes('deduction from payment')) {

return 'Deduction from payment';

}


return typeStr; // 如果没有匹配到关键词，返回原始清理后的字符串

}

return '';

}

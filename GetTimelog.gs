/**
 * Get Wrike Timelogs and output to spreadsheet / WrikeのTimelogを取得し、スプレッドシートに出力する
 */
function GetTimelog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('GetTimeLog');

  if (!sheet) {
    sheet = ss.insertSheet('GetTimeLog');
  }

  sheet.clear();

  const headers = ['TimelogId', 'taskId', 'taskName', 'taskLink', 'userId', 'User名', 'trackedDate', 'hours', 'comment', 'createdDate', 'updatedDate', 'approvalStatus', 'billingType', 'lockStatus', 'effortAllocation'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const apiUrl = scriptProperties.getProperty('api_url');
  const token = scriptProperties.getProperty('token');
  const options = {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  };

  // 1. Get Timelogs / タイムログを取得
  const timelogUrl = apiUrl + '/timelogs';
  const timelogResponse = UrlFetchApp.fetch(timelogUrl, options);
  const timelogData = JSON.parse(timelogResponse.getContentText()).data;

  if (!timelogData || timelogData.length === 0) {
    console.log('No timelogs found. / タイムログが見つかりませんでした。');
    return;
  }

  // 2. Get Contacts (Users) for mapping / マッピング用のコンタクト（ユーザー）を取得
  const contactsUrl = apiUrl + '/contacts';
  const contactsResponse = UrlFetchApp.fetch(contactsUrl, options);
  const contactsData = JSON.parse(contactsResponse.getContentText()).data;
  const userMap = {};
  contactsData.forEach(contact => {
    userMap[contact.id] = (contact.firstName || '') + ' ' + (contact.lastName || '');
  });

  // 3. Get unique task IDs and fetch task names and permalinks / ユニークなタスクIDを取得し、タスク名とパーマリンクを取得
  const taskIds = [...new Set(timelogData.map(log => log.taskId))];
  const taskMap = {};
  const taskLinkMap = {};
  const taskEffortAllocationMap = {};

  // Batch fetch tasks if possible, or fetch individually (Wrike API limits may apply) / 可能であればタスクを一括取得、または個別に取得（Wrike APIの制限が適用される場合があります）
  // Get multiple tasks at once: GET /tasks/{taskId1,taskId2,...} / 一度に複数のタスクを取得
  // Max tasks per request is usually around 100 / 1リクエストあたりの最大タスク数は通常100個程度
  const batchSize = 100;
  for (let i = 0; i < taskIds.length; i += batchSize) {
    const chunk = taskIds.slice(i, i + batchSize);
    const tasksUrl = apiUrl + '/tasks/' + chunk.join(',') + '?fields=' + encodeURIComponent('["effortAllocation"]');
    try {
      const tasksResponse = UrlFetchApp.fetch(tasksUrl, options);
      const tasksData = JSON.parse(tasksResponse.getContentText()).data;
      tasksData.forEach(task => {
        taskMap[task.id] = task.title;
        taskLinkMap[task.id] = task.permalink;
        taskEffortAllocationMap[task.id] = task.effortAllocation ? JSON.stringify(task.effortAllocation) : '';
      });
    } catch (e) {
      console.error('Error fetching tasks: / タスクの取得中にエラーが発生しました: ' + e.message);
    }
  }

  // 4. Process and write data / データの処理と書き込み
  const rows = timelogData.map(log => {
    return [
      log.id,
      log.taskId,
      taskMap[log.taskId] || '',
      taskLinkMap[log.taskId] || '',
      log.userId,
      userMap[log.userId] || '',
      log.trackedDate,
      log.hours,
      log.comment || '',
      log.createdDate,
      log.updatedDate,
      log.approvalStatus,
      log.billingType,
      log.lockStatus,
      taskEffortAllocationMap[log.taskId] || ''
    ];
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  SpreadsheetApp.getUi().alert('Timelog retrieval complete. / タイムログの取得が完了しました。');
}

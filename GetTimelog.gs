/**
 * WrikeのTimelogを取得し、スプレッドシートに出力する
 */
function GetTimelog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('GetTimeLog');

  if (!sheet) {
    sheet = ss.insertSheet('GetTimeLog');
  }

  sheet.clear();

  const headers = ['ID', 'taskId', 'taskName', 'userId', 'User 名', 'trackedDate', 'hours', 'comment', 'createdDate', 'updatedDate', 'approvalStatus', 'billingType', 'lockStatus'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const apiUrl = scriptProperties.getProperty('api_url');
  const token = scriptProperties.getProperty('token');
  const options = {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  };

  // 1. Get Timelogs
  const timelogUrl = apiUrl + '/timelogs';
  const timelogResponse = UrlFetchApp.fetch(timelogUrl, options);
  const timelogData = JSON.parse(timelogResponse.getContentText()).data;

  if (!timelogData || timelogData.length === 0) {
    console.log('タイムログが見つかりませんでした。');
    return;
  }

  // 2. Get Contacts (Users) for mapping
  const contactsUrl = apiUrl + '/contacts';
  const contactsResponse = UrlFetchApp.fetch(contactsUrl, options);
  const contactsData = JSON.parse(contactsResponse.getContentText()).data;
  const userMap = {};
  contactsData.forEach(contact => {
    userMap[contact.id] = (contact.firstName || '') + ' ' + (contact.lastName || '');
  });

  // 3. Get unique task IDs and fetch task names
  const taskIds = [...new Set(timelogData.map(log => log.taskId))];
  const taskMap = {};

  // Batch fetch tasks if possible, or fetch individually (Wrike API limits may apply)
  // Get multiple tasks at once: GET /tasks/{taskId1,taskId2,...}
  // Max tasks per request is usually around 100
  const batchSize = 100;
  for (let i = 0; i < taskIds.length; i += batchSize) {
    const chunk = taskIds.slice(i, i + batchSize);
    const tasksUrl = apiUrl + '/tasks/' + chunk.join(',');
    try {
      const tasksResponse = UrlFetchApp.fetch(tasksUrl, options);
      const tasksData = JSON.parse(tasksResponse.getContentText()).data;
      tasksData.forEach(task => {
        taskMap[task.id] = task.title;
      });
    } catch (e) {
      console.error('Error fetching tasks: ' + e.message);
    }
  }

  // 4. Process and write data
  const rows = timelogData.map(log => {
    return [
      log.id,
      log.taskId,
      taskMap[log.taskId] || '',
      log.userId,
      userMap[log.userId] || '',
      log.trackedDate,
      log.hours,
      log.comment || '',
      log.createdDate,
      log.updatedDate,
      log.approvalStatus,
      log.billingType,
      log.lockStatus
    ];
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  SpreadsheetApp.getUi().alert('タイムログの取得が完了しました。');
}

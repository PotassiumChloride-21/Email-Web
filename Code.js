// 邮件发送工具 - 主代码文件
// 文件名: Code.gs (不是script.gs，Google Apps Script默认用Code.gs)

function doGet() {
  try {
    return HtmlService
      .createTemplateFromFile('Index')
      .evaluate()
      .setTitle('邮件发送工具')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return HtmlService
      .createHtmlOutput('<h1>错误</h1><p>页面加载失败：' + error.message + '</p>')
      .setTitle('错误')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

// 包含HTML文件（正确的方法）
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// 检查授权状态
function checkAuthorization() {
  try {
    // 尝试访问Drive API
    DriveApp.getRootFolder();
    return {
      authorized: true,
      message: '已授权访问Google Drive',
      userEmail: Session.getActiveUser().getEmail()
    };
  } catch (error) {
    return {
      authorized: false,
      message: '需要授权访问Google Drive：' + error.message,
      authUrl: ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL).getAuthorizationUrl()
    };
  }
}

// 发送邮件
function sendCustomEmail(recipient, subject, body, cc, bcc, attachmentData) {
  try {
    // 参数处理
    cc = cc || '';
    bcc = bcc || '';
    attachmentData = attachmentData || [];
    
    // 验证收件人
    if (!recipient || !isValidEmail(recipient)) {
      throw new Error('收件人邮箱格式不正确');
    }
    
    // 验证邮件内容
    if (!subject || !subject.trim()) {
      throw new Error('邮件主题不能为空');
    }
    
    if (!body || !body.trim()) {
      throw new Error('邮件内容不能为空');
    }
    
    // 准备附件
    let attachments = [];
    let attachmentUrls = [];
    
    if (attachmentData.length > 0) {
      for (let i = 0; i < attachmentData.length; i++) {
        try {
          const fileInfo = attachmentData[i];
          if (fileInfo.id) {
            const file = DriveApp.getFileById(fileInfo.id);
            if (file) {
              attachments.push(file.getBlob());
              attachmentUrls.push({
                name: fileInfo.name || '附件',
                url: fileInfo.url || file.getUrl(),
                size: fileInfo.size || file.getSize()
              });
            }
          }
        } catch (e) {
          console.error('附件处理失败：', e.message);
        }
      }
    }
    
    // 添加附件信息到邮件正文
    let finalBody = body;
    if (attachmentUrls.length > 0) {
      finalBody += '\n\n---\n📎 附件列表：\n';
      attachmentUrls.forEach(item => {
        const sizeMB = (item.size / (1024 * 1024)).toFixed(2);
        finalBody += `• ${item.name} (${sizeMB} MB) - ${item.url}\n`;
      });
      finalBody += '---\n';
    }
    
    // 发送邮件
    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      body: finalBody,
      cc: cc || undefined,
      bcc: bcc || undefined,
      attachments: attachments
    });
    
    // 记录日志
    logEmailSent(recipient, subject, finalBody, attachments.length);
    
    return {
      success: true,
      message: '邮件发送成功！' + (attachments.length > 0 ? ` (包含${attachments.length}个附件)` : '')
    };
    
  } catch (error) {
    console.error('邮件发送失败：', error);
    return {
      success: false,
      message: '发送失败：' + error.message
    };
  }
}

// 上传文件到Google Drive
function uploadFileToDrive(fileName, base64Data, mimeType) {
  try {
    // 检查授权
    const authCheck = checkAuthorization();
    if (!authCheck.authorized) {
      throw new Error('未授权访问Google Drive');
    }
    
    // 检查文件大小
    const fileSize = (base64Data.length * 3) / 4;
    if (fileSize > 25 * 1024 * 1024) {
      return { 
        success: false, 
        message: '文件大小超过25MB限制' 
      };
    }
    
    // 创建文件夹
    const today = new Date();
    const folderName = '邮件附件_' + Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyyMMdd');
    let folder;
    
    const folders = DriveApp.getFoldersByName(folderName);
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }
    
    // 创建文件
    const decodedData = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decodedData, mimeType, fileName);
    const file = folder.createFile(blob);
    
    // 设置分享权限
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return {
      success: true,
      id: file.getId(),
      name: fileName,
      url: file.getUrl(),
      size: file.getSize(),
      mimeType: mimeType,
      message: '文件上传成功'
    };
    
  } catch (error) {
    console.error('文件上传失败：', error);
    return {
      success: false,
      message: '文件上传失败：' + error.message
    };
  }
}

// 模板管理
function saveTemplate(name, subject, body) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const templates = JSON.parse(userProperties.getProperty('emailTemplates') || '[]');
    
    templates.push({
      name: name,
      subject: subject,
      body: body,
      created: new Date().toISOString()
    });
    
    userProperties.setProperty('emailTemplates', JSON.stringify(templates));
    return { success: true, message: '模板保存成功' };
  } catch (error) {
    return { success: false, message: '保存失败：' + error.message };
  }
}

function getTemplates() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    return JSON.parse(userProperties.getProperty('emailTemplates') || '[]');
  } catch (error) {
    console.error('获取模板失败：', error);
    return [];
  }
}

function deleteTemplate(index) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const templates = JSON.parse(userProperties.getProperty('emailTemplates') || '[]');
    
    if (index >= 0 && index < templates.length) {
      templates.splice(index, 1);
      userProperties.setProperty('emailTemplates', JSON.stringify(templates));
      return { success: true, message: '模板删除成功' };
    }
    
    return { success: false, message: '模板不存在' };
  } catch (error) {
    return { success: false, message: '删除失败：' + error.message };
  }
}

// 邮件记录
function getRecentEmails(maxResults) {
  try {
    maxResults = maxResults || 5;
    const threads = GmailApp.search('from:me', 0, maxResults);
    const emails = [];
    
    for (let i = 0; i < threads.length && emails.length < maxResults; i++) {
      const messages = threads[i].getMessages();
      for (let j = 0; j < messages.length && emails.length < maxResults; j++) {
        const message = messages[j];
        emails.push({
          to: message.getTo(),
          subject: message.getSubject(),
          date: message.getDate().toISOString(),
          body: message.getPlainBody().substring(0, 100) + '...'
        });
      }
    }
    
    return emails;
  } catch (error) {
    console.error('获取邮件记录失败：', error);
    return getEmailLogs(maxResults);
  }
}

function logEmailSent(recipient, subject, body, attachmentCount) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const logs = JSON.parse(scriptProperties.getProperty('emailLogs') || '[]');
    
    logs.unshift({
      to: recipient,
      subject: subject,
      body: body.substring(0, 200),
      attachments: attachmentCount,
      timestamp: new Date().toISOString()
    });
    
    // 只保留最近50条
    if (logs.length > 50) {
      logs.length = 50;
    }
    
    scriptProperties.setProperty('emailLogs', JSON.stringify(logs));
  } catch (e) {
    console.error('记录日志失败：', e);
  }
}

function getEmailLogs(maxResults) {
  try {
    maxResults = maxResults || 10;
    const scriptProperties = PropertiesService.getScriptProperties();
    const logs = JSON.parse(scriptProperties.getProperty('emailLogs') || '[]');
    
    return logs.slice(0, maxResults).map(log => ({
      to: log.to,
      subject: log.subject,
      body: log.body,
      attachments: log.attachments,
      date: new Date(log.timestamp).toLocaleString()
    }));
  } catch (e) {
    console.error('获取日志失败：', e);
    return [];
  }
}

// 实用函数
function isValidEmail(email) {
  if (!email) return false;
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

function getMaxFileSize() {
  return {
    maxSize: 25 * 1024 * 1024, // 25MB
    maxTotalSize: 50 * 1024 * 1024 // 50MB
  };
}

function cleanupTempFiles() {
  try {
    const oneDayAgo = new Date();
    oneDayAgo.setDate(oneDayAgo.getDate() - 1);
    
    let deletedCount = 0;
    const folders = DriveApp.getFolders();
    
    while (folders.hasNext()) {
      const folder = folders.next();
      if (folder.getName().startsWith('邮件附件_')) {
        const files = folder.getFiles();
        
        while (files.hasNext()) {
          const file = files.next();
          if (file.getDateCreated() < oneDayAgo) {
            try {
              file.setTrashed(true);
              deletedCount++;
            } catch (e) {
              console.log('无法删除文件：', e.message);
            }
          }
        }
      }
    }
    
    return { 
      success: true, 
      message: '清理完成，删除了 ' + deletedCount + ' 个文件' 
    };
  } catch (error) {
    return { 
      success: false, 
      message: '清理失败：' + error.message 
    };
  }
}

// 测试函数
function testConnection() {
  return {
    success: true,
    message: '服务器连接正常',
    timestamp: new Date().toISOString(),
    userEmail: Session.getActiveUser().getEmail()
  };
}
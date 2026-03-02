/******************************
 * LINE + Google Calendar Bot + Gemini
 * by I3AYY & ChatGPT Assistant
 ******************************/

// === CONFIGURATION ===

// --- LINE ---
const LINE_CHANNEL_ACCESS_TOKEN = 'XXXXX'; // ใส่ Access Token ของ LINE Bot ของคุณ

const LINE_USER_ID = 'XXXXX';              // ใส่ User ID หรือ Group ID ของ LINE ที่ต้องการส่งแจ้งเตือน

// --- Google Calendar ---
const CALENDAR_ID = 'XXXXX';           // ใส่ Calendar ID ที่ต้องการใช้งาน (หรือ 'primary')

// --- Google Gemini AI ---
const GEMINI_API_KEY = 'XXXXX';            // ใส่ API Key ของ Gemini ที่คุณสร้างขึ้น

// === Updated doPost with postback handling ===
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const event = data.events && data.events[0];
    if (!event) return ContentService.createTextOutput('OK');

    // ---------- 1) ใน doPost: ก่อน process message ให้จับ postback (วางหลัง const event = ... )
    if (event.postback && event.postback.data) {
      try {
        return handlePostback(event.postback, event.replyToken);
      } catch (err) {
        console.error('handlePostback error:', err);
        // fallthrough to normal processing
      }
    }

    // Support shortcuts flex via message "คำสั่งลัด"
    const userMessage = String((event.message && event.message.text) || '').trim();
    const replyToken = event.replyToken;
    const quoteToken = event.message && event.message.quoteToken ? event.message.quoteToken : null; // ดึง quoteToken ถ้ามี (เฉพาะ text/sticker)

    // show loading in 1:1 chats (best-effort), reduced to 10 seconds to avoid long hangs
    try {
      const source = event.source || {};
      if (source.type === 'user' && source.userId) {
        displayLoading(source.userId, 10);
      }
    } catch (ignoreLoadingErr) { console.warn('displayLoading warning:', ignoreLoadingErr); }

    if (!userMessage) {
      const flex = makeActionBubble('info', 'วิธีใช้', 'พิมพ์เป็นประโยคธรรมชาติได้เลย เช่น: "พรุ่งนี้มีอะไรบ้าง", "เพิ่มนัด 30 สิงหา เรื่อง โอน 3500"');
      return replyFlexWithQuick(replyToken, flex); // now this sends flex only (no quick replies)
    }

    // If user requests shortcuts
    if (/^คำสั่งลัด$/i.test(userMessage) || /^shortcuts$/i.test(userMessage)) {
      const flex = makeShortcutsFlex();
      return replyFlexWithQuick(replyToken, flex);
    }

    // เรียก Gemini (parser)
    let structuredInfo = callGemini(userMessage);

    // fallback local parse
    if (!structuredInfo || structuredInfo.action === 'unknown') {
      const local = tryLocalThaiParse(userMessage);
      if (local && local.action && local.action !== 'unknown') structuredInfo = local;
    }

    // resolve parsed tokens (ถ้ามี)
    if (structuredInfo && structuredInfo.startTime) {
      const resolved = resolveParsedDateString(structuredInfo.startTime, structuredInfo);
      if (resolved) structuredInfo._parsedStart = resolved;
    }
    if (structuredInfo && structuredInfo.endTime) {
      const resolvedEnd = resolveParsedDateString(structuredInfo.endTime, structuredInfo);
      if (resolvedEnd) structuredInfo._parsedEnd = resolvedEnd;
    }
    if (structuredInfo && structuredInfo.newStartTime) {
      const resolvedNew = resolveParsedDateString(structuredInfo.newStartTime, structuredInfo);
      if (resolvedNew) structuredInfo._parsedNewStart = resolvedNew;
    }
    if (structuredInfo && structuredInfo.newEndTime) {
      const resolvedNewEnd = resolveParsedDateString(structuredInfo.newEndTime, structuredInfo);
      if (resolvedNewEnd) structuredInfo._parsedNewEnd = resolvedNewEnd;
    }

    // time specified flag & wantsAllDay flag
    structuredInfo = structuredInfo || {};
    structuredInfo._timeSpecified = detectTimeMention(userMessage) || false;
    structuredInfo._wantsAllDay = /ทั้งวัน|เป็นทั้งวัน|all[- ]?day/i.test(userMessage);

    // recurring fallback (เช่น "ทุกวันที่ X")
    if ((!structuredInfo || structuredInfo.action === 'unknown') && detectRecurringInText(userMessage)) {
      const rec = detectRecurringInText(userMessage);
      if (rec && rec.type === 'monthly') {
        structuredInfo = { action: 'create', title: rec.title || ('นัดประจำเดือนวันที่ ' + rec.day), dayOfMonth: rec.day };
      }
    }

    // ถ้ายังเป็น unknown => ส่งให้ Gemini ตอบคำถามทั่วไป (short answer)
    if (!structuredInfo || structuredInfo.action === 'unknown') {
      const shortAnswer = askGeminiShort(userMessage);
      if (shortAnswer) {
        return replyTextWithQuick(replyToken, shortAnswer, quoteToken); // ส่ง quoteToken สำหรับ text reply (เฉพาะ non-calendar)
      } else {
        const flex = makeActionBubble('info', 'ไม่เข้าใจคำสั่ง', 'ลองพิมพ์แบบตัวอย่าง:\n• เพิ่มนัด 30 สิงหา เรื่อง โอน 3500\n• ปรับเวลา เป็นทั้งวัน เรื่อง โอน 3500 วันที่ 30 สิงหา\n• ลบนัด ประชุมโปรเจค\nพิมพ์ "คำสั่งลัด" เพื่อเรียกเมนูคำสั่งลัด');
        return replyFlexWithQuick(replyToken, flex); // flex ไม่ quote
      }
    }

    // หากเป็นคำสั่ง Calendar (create/delete/update/list) ให้ dispatch ตามเดิม
    const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

    switch (structuredInfo.action) {
      case 'create':
        handleCreateEvent(calendar, replyToken, structuredInfo, userMessage);
        break;
      case 'delete':
        handleDeleteEvent(calendar, replyToken, structuredInfo);
        break;
      case 'update':
        handleUpdateEvent(calendar, replyToken, structuredInfo, userMessage);
        break;
      case 'list':
        handleListEvents(calendar, replyToken, structuredInfo, userMessage);
        break;
      default:
        const flex = makeActionBubble('info', 'ไม่แน่ใจ', `ฉันเห็นว่าคุณต้องการ "${structuredInfo.title || '…'}" แต่ไม่แน่ใจว่าจะทำอะไร`);
        replyFlexWithQuick(replyToken, flex);
    }

    return ContentService.createTextOutput('OK');
  } catch (err) {
    console.error('doPost error:', err);
    try {
      const data = JSON.parse(e.postData.contents);
      const replyToken = data.events[0].replyToken;
      const flex = makeActionBubble('info', 'ข้อผิดพลาด', 'เกิดข้อผิดพลาดในระบบ ลองอีกครั้งภายหลัง');
      replyFlexWithQuick(replyToken, flex);
    } catch (_) {}
    return ContentService.createTextOutput('OK');
  }
}

// ---------- 2) ฟังก์ชันจัดการ postback (ใหม่)
function handlePostback(postback, replyToken) {
  const dataStr = postback.data || '';
  let payload;
  try {
    payload = JSON.parse(dataStr);
  } catch (e) {
    console.warn('Postback not JSON:', dataStr);
    // ถ้าไม่ใช่ JSON แบบที่เราคาดไว้ ให้แจ้งผู้ใช้เล็กน้อย
    return replyTextWithQuick(replyToken, 'คำสั่งยืนยันไม่ถูกต้อง (โปรดลองอีกครั้ง)');
  }

  if (payload.cmd === 'confirm_delete' && payload.key) {
    const properties = PropertiesService.getScriptProperties();
    const storedIds = properties.getProperty(payload.key);
    if (storedIds) {
      try {
        const ids = JSON.parse(storedIds);
        if (Array.isArray(ids) && ids.length) {
          performDeleteByIds(ids, replyToken);
        } else {
          replyTextWithQuick(replyToken, 'ไม่มีรายการที่จะลบ');
        }
      } catch (e) {
        console.error('Parse stored ids error:', e);
        replyTextWithQuick(replyToken, 'เกิดข้อผิดพลาดในการลบ');
      } finally {
        properties.deleteProperty(payload.key);
      }
    } else {
      replyTextWithQuick(replyToken, 'คำสั่งยืนยันหมดอายุหรือไม่ถูกต้อง');
    }
    return ContentService.createTextOutput('OK');
  }

  if (payload.cmd === 'cancel') {
    const flex = makeActionBubble('delete', 'ยกเลิก', 'ยกเลิกการลบเรียบร้อยแล้ว');
    replyFlexWithQuick(replyToken, flex);
    return ContentService.createTextOutput('OK');
  }

  // ไม่รู้จักคำสั่ง
  return replyTextWithQuick(replyToken, 'คำสั่ง postback ไม่รู้จัก');
}

// ---------- 3) ฟังก์ชันส่ง confirmation (ใหม่)
function sendDeleteConfirmation(replyToken, candidates) {
  // สร้างบรรทัดข้อความสำหรับ body ของ Flex
  const maxShow = 10;
  const lines = candidates.slice(0, maxShow).map(c => {
    const dt = new Date(c.startTimeISO);
    const timeStr = c.isAllDay ? `${formatDate(dt)} (ทั้งวัน)` : `${formatDate(dt)} ${formatTime(dt)}`;
    return `• ${timeStr} : ${c.title}`;
  });

  if (candidates.length > maxShow) lines.push(`... และอีก ${candidates.length - maxShow} รายการ`);

  // สร้าง postback payload (stringify)
  const properties = PropertiesService.getScriptProperties();
  const deleteKey = 'del_' + Utilities.getUuid().slice(0, 8); // key สั้น unique
  properties.setProperty(deleteKey, JSON.stringify(candidates.map(c => c.id)));
  const confirmPayload = JSON.stringify({ cmd: 'confirm_delete', key: deleteKey });
  const cancelPayload = JSON.stringify({ cmd: 'cancel' });

  const bubble = {
    type: 'flex',
    altText: 'ยืนยันการลบนัด',
    contents: {
      type: 'bubble',
      header: {
        type: 'box',
        layout: 'vertical',
        contents: [{ type: 'text', text: 'ยืนยันการลบ', weight: 'bold', size: 'lg', color: '#ffffff' }],
        backgroundColor: '#dc3545',
        paddingAll: 'md'
      },
      body: {
        type: 'box',
        layout: 'vertical',
        spacing: 'sm',
        contents: [
          { type: 'text', text: 'คุณกำลังจะลบรายการดังต่อไปนี้:', size: 'sm', wrap: true },
          { type: 'text', text: lines.join('\n'), size: 'sm', wrap: true },
          { type: 'text', text: '\nต้องการลบรายการทั้งหมดเหล่านี้หรือไม่?', size: 'sm', wrap: true }
        ]
      },
      footer: {
        type: 'box',
        layout: 'horizontal',
        spacing: 'sm',
        contents: [
          {
            type: 'button',
            style: 'primary',
            color: '#dc3545',
            action: { type: 'postback', label: 'ยืนยันลบ', data: confirmPayload }
          },
          {
            type: 'button',
            style: 'secondary',
            action: { type: 'postback', label: 'ยกเลิก', data: cancelPayload }
          }
        ]
      }
    }
  };

  // ส่ง flex
  replyFlexWithQuick(replyToken, bubble);
}

// ---------- 4) ฟังก์ชันลบตาม id ที่ยืนยันแล้ว (ใหม่)
function performDeleteByIds(ids, replyToken) {
  const lines = [];
  let firstDate = null;
  let deletedCount = 0;

  ids.forEach(id => {
    try {
      // CalendarApp.getEventById ใช้ id ที่ event.getId() ให้มา
      const ev = CalendarApp.getEventById(id);
      if (ev) {
        const start = ev.getStartTime();
        if (!firstDate) firstDate = start;
        const isAll = ev.isAllDayEvent && ev.isAllDayEvent();
        const line = isAll
          ? `• [ทั้งวัน] ${formatDate(start)} : ${ev.getTitle()}`
          : `• ${formatDate(start)} ${formatTime(start)} - ${formatTime(ev.getEndTime())} : ${ev.getTitle()}`;
        lines.push(line);
        ev.deleteEvent();
        deletedCount++;
      } else {
        // หากไม่พบ event โดย id อาจเกิดจาก id หมดอายุหรือผิดฟอร์แมต
        lines.push(`• ไม่พบ event id: ${id}`);
      }
    } catch (err) {
      console.error('performDeleteByIds error for id', id, err);
      lines.push(`• ไม่สามารถลบ event id: ${id} (ข้อผิดพลาด)`);
    }
  });

  const title = deletedCount ? 'ลบเรียบร้อย' : 'ลบไม่สำเร็จ';
  const body = deletedCount ? `ลบงานแล้ว ${deletedCount} รายการ:\n` + lines.join('\n') : 'ไม่พบเหตุการณ์ที่จะลบหรือเกิดข้อผิดพลาด\n' + lines.join('\n');
  const flex = makeActionBubble('delete', title, body, firstDate);
  replyFlexWithQuick(replyToken, flex);
}

// ==============================================
// NEW FUNCTION: displayLoading เรียก LINE Messaging API: POST /v2/bot/chat/loading/start
function displayLoading(chatId, seconds) {
  try {
    if (!chatId) return null;
    const sec = Math.max(5, Math.min(60, parseInt(seconds, 10) || 5)); // clamp 5..60
    const url = 'https://api.line.me/v2/bot/chat/loading/start';
    const payload = {
      chatId: chatId,
      loadingSeconds: sec
    };
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + LINE_CHANNEL_ACCESS_TOKEN },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    const resp = UrlFetchApp.fetch(url, options);
    const code = resp.getResponseCode();
    if (code >= 200 && code < 300) {
      return true;
    } else {
      console.warn('displayLoading failed:', code, resp.getContentText());
      return false;
    }
  } catch (err) {
    console.error('displayLoading error:', err);
    return false;
  }
}

// === เพิ่มฟังก์ชัน: askGeminiShort ===
function askGeminiShort(userText) {
  try {
    const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-lite:generateContent?key=${GEMINI_API_KEY}`;

    const prompt = `
คุณคือ LINE AI Agent Model: Gemini 2.5 Flash-Lite ที่ทำงานอยู่ในระบบ LINE OA
ตอบคำถามทั่วไปได้ทุกเรื่อง ตามความสามารถของ Gemini 2.5 Flash-Lite
ตอบเป็นภาษาเดียวกับผู้ใช้ แบบสุภาพ กระชับ และเข้าใจง่าย ไม่เกิน 5 บรรทัด

ข้อความผู้ใช้: "${userText}"
`;

    const payload = {
      contents: [{ parts: [{ text: prompt }] }],
      safetySettings: [
        { category: "HARM_CATEGORY_HARASSMENT", threshold: "BLOCK_NONE" },
        { category: "HARM_CATEGORY_HATE_SPEECH", threshold: "BLOCK_NONE" },
        { category: "HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold: "BLOCK_NONE" },
        { category: "HARM_CATEGORY_DANGEROUS_CONTENT", threshold: "BLOCK_NONE" }
      ]
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(apiUrl, options);
    const code = response.getResponseCode();
    const text = response.getContentText();
    if (code !== 200) {
      console.error('Gemini short reply error:', code, text);
      return null;
    }

    const result = JSON.parse(text);
    const candidateText = result.candidates && result.candidates[0] && result.candidates[0].content && result.candidates[0].content.parts && result.candidates[0].content.parts[0] && result.candidates[0].content.parts[0].text;
    if (!candidateText) return null;

    let answer = candidateText.replace(/^```(?:\w+)?\n?|```$/g, '').trim();
    const lines = answer.split(/\r?\n/).filter(Boolean);
    if (lines.length > 3) answer = lines.slice(0,3).join('\n');

    return answer;
  } catch (err) {
    console.error('askGeminiShort error:', err);
    return null;
  }
}

// === reply helpers (quick replies removed) ===
function replyTextWithQuick(replyToken, messageText, quoteToken = null) {
  try {
    const url = 'https://api.line.me/v2/bot/message/reply';
    const message = { type: 'text', text: messageText };
    if (quoteToken) {
      message.quoteToken = quoteToken; // เพิ่ม quoteToken สำหรับ text reply ถ้ามี
    }
    const payload = { replyToken: replyToken, messages: [message] };
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + LINE_CHANNEL_ACCESS_TOKEN },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    return ContentService.createTextOutput('OK');
  } catch (err) {
    console.error('replyTextWithQuick error:', err);
    return ContentService.createTextOutput('OK');
  }
}

function replyFlexWithQuick(replyToken, flexMessage) {
  // Renamed behavior: this now replies only the flexMessage (no quick replies)
  try {
    const url = 'https://api.line.me/v2/bot/message/reply';
    const payload = { replyToken: replyToken, messages: [flexMessage] };
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + LINE_CHANNEL_ACCESS_TOKEN },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    return ContentService.createTextOutput('OK');
  } catch (err) {
    console.error('replyFlexWithQuick error:', err);
    return ContentService.createTextOutput('OK');
  }
}

// === EVENT HANDLERS ===
function handleCreateEvent(calendar, replyToken, info, originalText) {
  const { title, description, location } = info;

  // Resolve start/end (prefer parsed token)
  let startDate = info._parsedStart || (info.startTime ? resolveParsedDateString(info.startTime, info) : null);
  let endDate = info._parsedEnd || (info.endTime ? resolveParsedDateString(info.endTime, info) : null);

  // 1) ตรวจหา recurring ที่ชัดเจนจากข้อความ (เฉพาะกรณีที่ detectRecurringInText ระบุ monthly)
  // --- recurring monthly (with series tag)
  const recurringFromText = detectRecurringInText(originalText);
  if (recurringFromText && recurringFromText.type === 'monthly') {
    const day = recurringFromText.day;
    if (!day || day < 1 || day > 31) {
      const flex = makeActionBubble('delete', 'รูปแบบไม่ถูกต้อง', `ไม่พบวันที่สำหรับสร้างแบบประจำเดือน (เช่น "ทุกวันที่ 27")`);
      return replyFlexWithQuick(replyToken, flex);
    }

    // สร้าง seriesId เดียวกันสำหรับทั้งชุด เพื่อใช้ตอนลบหรือเช็ก
    const seriesId = `series:monthly:${day}:${Date.now()}`;
    const createdDates = [];
    const now = new Date();
    for (let i = 0; i < 12; i++) {
      const dt = new Date(now.getFullYear(), now.getMonth() + i, day, 0, 0, 0);
      if (dt.getDate() !== day) continue;
      const descWithSeries = (description || '') + (description ? '\n\n' : '') + seriesId;
      calendar.createAllDayEvent(title || ('นัดประจำเดือนวันที่ ' + day), dt, { description: descWithSeries, location: location || '' });
      createdDates.push(formatDate(dt));
    }
    if (createdDates.length === 0) {
      const flex = makeActionBubble('delete', 'ไม่สามารถสร้างได้', 'ไม่สามารถสร้างนัดประจำเดือนสำหรับวันนั้นๆ ได้ในช่วง 12 เดือนถัดไป');
      return replyFlexWithQuick(replyToken, flex);
    }
    const flex = makeActionBubble('create', 'สร้างนัดประจำเดือนแล้ว', `เรื่อง: ${title || ('นัดประจำเดือนวันที่ ' + day)}\nสร้างล่วงหน้า ${createdDates.length} งวด:\n• ${createdDates.join('\n• ')}\n\n(Series ID: ${seriesId})`, null);
    return replyFlexWithQuick(replyToken, flex);
  }

  // 2) หาก parser ให้ dayOfMonth แต่ไม่มี parsedStart ให้พยายามตีความเป็น 'วันเฉพาะ' แทน 'ทุกเดือน'
  if (!startDate && info.dayOfMonth) {
    // ถ้าข้อความมีเดือนที่ชัดเจน ให้พยายาม parse วันที่เฉพาะ
    const parsedFromText = parseDateFromThaiText(originalText);
    if (parsedFromText) {
      startDate = parsedFromText;
    } else {
      // ถ้าไม่มี month ระบุ: สร้างเป็น occurrence ใกล้ที่สุดของเลขวันนั้น (this month หรือ next month)
      const now = new Date();
      const candidate = new Date(now.getFullYear(), now.getMonth(), info.dayOfMonth, 0,0,0);
      const todayAtStart = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0,0,0);
      if (candidate < todayAtStart) {
        candidate.setMonth(candidate.getMonth() + 1);
      }
      if (candidate.getDate() === info.dayOfMonth) {
        startDate = candidate;
      }
    }
  }

  // 3) ถ้ายังไม่มี startDate ให้ลอง parse จากข้อความโดยตรง
  if (!startDate) {
    const localDate = parseDateFromThaiText(originalText);
    if (localDate) startDate = localDate;
  }

  // 4) ถ้ายังไม่มีเลย => สร้างเป็นงานทั้งวันวันนี้ (fallback เดิม)
  if (!startDate) {
    const today = new Date();
    calendar.createAllDayEvent(title, today, { description: description || '', location: location || '' });
    const flex = makeActionBubble('create', 'สร้างงานทั้งวันแล้ว', `เรื่อง: ${title}\nวันที่: ${formatDate(today)}${location ? `\nสถานที่: ${location}` : ''}${description ? `\nรายละเอียด: ${description}` : ''}`, today);
    return replyFlexWithQuick(replyToken, flex);
  }

  // determine all-day vs timed
  const timeSpecified = !!info._timeSpecified;
  let ev;
  if (!timeSpecified || info._wantsAllDay) {
    const d = new Date(startDate); d.setHours(0,0,0,0);
    ev = calendar.createAllDayEvent(title, d, { description: description || '', location: location || '' });
  } else {
    if (endDate) {
      ev = calendar.createEvent(title, startDate, endDate, { description: description || '', location: location || '' });
    } else {
      const defaultEnd = new Date(startDate.getTime() + 60 * 60 * 1000);
      ev = calendar.createEvent(title, startDate, defaultEnd, { description: description || '', location: location || '' });
    }
  }

  const flex = makeCreateConfirmForEvent(ev, title, description, location);
  return replyFlexWithQuick(replyToken, flex);
}

function handleDeleteEvent(calendar, replyToken, info) {
  const { title, dayOfMonth, isRecurring } = info;
  const today = new Date();
  const nextYear = new Date(today.getTime() + 365 * 24 * 60 * 60 * 1000);

  // ดึง events ภายใน 1 ปีข้างหน้า
  const events = calendar.getEvents(today, nextYear);
  const titleLower = (title || '').toLowerCase();

  // หา candidates โดย title (เท่ากับก่อน) ถ้าไม่พบให้ใช้ includes
  let candidates = events.filter(e => e.getTitle().toLowerCase() === titleLower);
  if (candidates.length === 0) candidates = events.filter(e => e.getTitle().toLowerCase().includes(titleLower));

  if (!candidates.length) {
    const flex = makeActionBubble('delete', 'ไม่พบงาน', `ไม่พบงานชื่อใกล้เคียง "${title}" ใน 365 วันข้างหน้า`);
    return replyFlexWithQuick(replyToken, flex);
  }

  // เตรียมรูปแบบ candidate objects (รวม id สำหรับการลบที่แน่นอน)
  const prepared = candidates.map(ev => {
    return {
      id: ev.getId ? ev.getId() : null,
      title: ev.getTitle(),
      startTimeISO: ev.getStartTime().toISOString(),
      isAllDay: (ev.isAllDayEvent && ev.isAllDayEvent()) || false,
      description: ev.getDescription ? ev.getDescription() : ''
    };
  });

  // ถ้าผู้ใช้สั่งลบ recurring หรือ detect dayOfMonth ให้ filter เฉพาะที่ตรงกับ dayOfMonth หรือ series tag
  let toDelete = [];
  if (isRecurring || dayOfMonth) {
    const day = dayOfMonth || (info._parsedStart && info._parsedStart.getDate());
    // หา events ที่มี series tag (ถ้ามี)
    prepared.forEach(p => {
      if (p.description && p.description.indexOf(`series:monthly:${day}:`) !== -1) toDelete.push(p);
    });
    // fallback: ถ้าไม่มี tag ให้เลือกเฉพาะ events ที่ start date day === day
    if (toDelete.length === 0) {
      prepared.forEach(p => {
        const d = new Date(p.startTimeISO);
        if (d.getDate() === parseInt(day, 10)) toDelete.push(p);
      });
    }

    if (toDelete.length === 0) {
      const flex = makeActionBubble('delete', 'ไม่พบงานแบบประจำ', `ไม่พบงานแบบประจำชื่อ "${title}" ที่ตรงกับวันที่ ${day} ใน 12 เดือนข้างหน้า`);
      return replyFlexWithQuick(replyToken, flex);
    }
  } else {
    // กรณีลบปกติ: ให้ยืนยันก่อนถ้ามีหลายรายการ
    toDelete = prepared;
  }

  // ถ้าเหลือแค่ 1 รายการและไม่ได้เป็น recurring -> ลบทันที (หรือคุณจะเปลี่ยนให้ confirm เสมอก็ได้)
  if (toDelete.length === 1 && !isRecurring && !dayOfMonth) {
    if (toDelete[0].id) {
      performDeleteByIds([toDelete[0].id], replyToken);
    } else {
      // ถ้าไม่มี id (แปลก) -> ลบโดยตรงจาก candidates (fallback)
      candidates[0].deleteEvent();
      const start = candidates[0].getStartTime();
      const flex = makeActionBubble('delete', 'ลบเรียบร้อย', `ลบงานแล้ว 1 รายการ:\n• ${formatDate(start)} ${formatTime(start)} : ${candidates[0].getTitle()}`, start);
      replyFlexWithQuick(replyToken, flex);
    }
    return;
  }

  // ถ้ามีหลายรายการ -> ส่ง confirmation flex (ผู้ใช้กดยืนยันแล้วจะส่ง postback ที่มี ids)
  sendDeleteConfirmation(replyToken, toDelete);
}

function handleUpdateEvent(calendar, replyToken, info, originalText) {
  const { title, newTitle, newStartTime, newEndTime, description, location } = info;
  const today = new Date();
  const nextYear = new Date(today.getTime() + 365 * 24 * 60 * 60 * 1000);
  const events = calendar.getEvents(today, nextYear);
  const exact = events.filter(e => title && e.getTitle().toLowerCase() === title.toLowerCase());
  const matches = exact.length ? exact : (title ? events.filter(e => e.getTitle().toLowerCase().includes(title.toLowerCase())) : []);

  if (!matches.length) {
    const flex = makeActionBubble('delete', 'ไม่พบงาน', `ไม่พบนัดที่จะอัปเดต (ค้นหา "${title || '—'}")`);
    return replyFlexWithQuick(replyToken, flex);
  }

  const wantsAllDay = !!info._wantsAllDay || /ทั้งวัน|เป็นทั้งวัน|all[- ]?day/i.test(originalText || '');
  const updatedLines = [];
  let firstDate = null;

  matches.forEach(ev => {
    if (!firstDate) firstDate = ev.getStartTime();

    if (wantsAllDay) {
      const evStart = ev.getStartTime();
      const allDayDate = new Date(evStart.getFullYear(), evStart.getMonth(), evStart.getDate(), 0,0,0);
      const newTitleToUse = newTitle || ev.getTitle();
      const newDesc = (typeof description === 'string') ? description : ev.getDescription();
      const newLoc = (typeof location === 'string') ? location : ev.getLocation();
      calendar.createAllDayEvent(newTitleToUse, allDayDate, { description: newDesc || '', location: newLoc || '' });
      ev.deleteEvent();
      updatedLines.push(`• [ทั้งวัน] ${formatDate(allDayDate)} : ${newTitleToUse}`);
      return;
    }

    if (newTitle) ev.setTitle(newTitle);

    if (info._parsedNewStart) {
      const s = info._parsedNewStart;
      if (!isNaN(s.getTime())) {
        if (info._parsedNewEnd) {
          const e = info._parsedNewEnd;
          if (!isNaN(e.getTime())) ev.setTime(s, e);
        } else {
          const oldStart = ev.getStartTime();
          const oldEnd = ev.getEndTime();
          const durMs = Math.max(oldEnd.getTime() - oldStart.getTime(), 30 * 60 * 1000);
          ev.setTime(s, new Date(s.getTime() + durMs));
        }
      }
    }

    if (typeof description === 'string') ev.setDescription(description);
    if (typeof location === 'string') ev.setLocation(location);

    const start = ev.getStartTime();
    const end = ev.getEndTime();
    const line = ev.isAllDayEvent()
      ? `• [ทั้งวัน] ${formatDate(start)} : ${ev.getTitle()}`
      : `• ${formatDate(start)} ${formatTime(start)} - ${formatTime(end)} : ${ev.getTitle()}`;
    updatedLines.push(line);
  });

  const flex = makeActionBubble('update', 'อัปเดตเรียบร้อย', `อัปเดตงานสำเร็จ ${matches.length} รายการ:\n` + updatedLines.join('\n'), firstDate);
  return replyFlexWithQuick(replyToken, flex);
}

function handleListEvents(calendar, replyToken, info, originalText) {
  if (/สัปดาห์หน้า/.test(originalText || '')) {
    const nextMonday = getNextWeekMonday(new Date());
    const bubbles = [];
    for (let i = 0; i < 7; i++) {
      const d = new Date(nextMonday);
      d.setDate(nextMonday.getDate() + i);
      const s = new Date(d); s.setHours(0,0,0,0);
      const e = new Date(d); e.setHours(23,59,59,999);
      const events = calendar.getEvents(s, e);
      bubbles.push(buildDayBubbleContent(d, events, null, 'list-day'));
    }
    const flex = { type: 'flex', altText: 'ตารางสัปดาห์หน้า', contents: { type: 'carousel', contents: bubbles } };
    return replyFlexWithQuick(replyToken, flex);
  }

  let targetDate = info._parsedStart || (info.startTime ? resolveParsedDateString(info.startTime, info) : null);
  if (!targetDate) targetDate = new Date();
  if (isNaN(targetDate.getTime())) {
    const flex = makeActionBubble('info', 'วันที่ไม่ถูกต้อง', `รูปแบบวันที่ไม่ถูกต้อง (${info.startTime})`);
    return replyFlexWithQuick(replyToken, flex);
  }

  const startTime = new Date(targetDate); startTime.setHours(0,0,0,0);
  const endTime = new Date(targetDate); endTime.setHours(23,59,59,999);
  const events = calendar.getEvents(startTime, endTime);

  // always build bubble; if empty it'll show "ไม่มีงาน"
  const bubble = buildDayBubbleContent(targetDate, events, null, 'list-day');
  const flex = { type: 'flex', altText: `ตารางวันที่ ${formatDate(targetDate)}`, contents: bubble };
  return replyFlexWithQuick(replyToken, flex);
}

// === NOTIFICATIONS / TRIGGERS ===
// Modified: always send flex (even if no events)
function sendTomorrowReminder() {
  try {
    const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
    const now = new Date();
    const tomorrow = new Date(now); tomorrow.setDate(now.getDate() + 1); tomorrow.setHours(0,0,0,0);
    const start = new Date(tomorrow);
    const end = new Date(tomorrow); end.setHours(23,59,59,999);
    const events = calendar.getEvents(start, end);
    const bubble = buildDayBubbleContent(tomorrow, events, 'ตารางสำหรับวันพรุ่งนี้', 'list-day');
    const flex = { type: 'flex', altText: 'ตารางวันพรุ่งนี้', contents: bubble };
    pushFlexToLine(flex);
  } catch (err) {
    console.error('sendTomorrowReminder error:', err);
  }
}

function sendAfternoonReminder() {
  try {
    const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
    const now = new Date();
    const start = new Date(now); start.setHours(12,0,0,0);
    const end = new Date(now); end.setHours(23,59,59,999);
    const events = calendar.getEvents(start, end);
    const bubble = buildDayBubbleContent(now, events, 'ตารางช่วงบ่ายวันนี้', 'list-day');
    const flex = { type: 'flex', altText: 'ตารางช่วงบ่ายวันนี้', contents: bubble };
    pushFlexToLine(flex);
  } catch (err) {
    console.error('sendAfternoonReminder error:', err);
  }
}

function sendMorningReminder() {
  try {
    const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
    const now = new Date();
    const today = new Date(now); today.setHours(0,0,0,0);

    const dates = [
      { date: new Date(today), title: 'ตารางงานวันนี้' },
      { date: new Date(today.getFullYear(), today.getMonth(), today.getDate() + 1), title: 'ตารางงานพรุ่งนี้' },
      { date: new Date(today.getFullYear(), today.getMonth(), today.getDate() + 2), title: 'ตารางงานมะรืนนี้' }
    ];

    const bubbles = [];

    dates.forEach(item => {
      const d = item.date;
      const s = new Date(d); s.setHours(0,0,0,0);
      const e = new Date(d); e.setHours(23,59,59,999);
      const events = calendar.getEvents(s, e);
      const bubble = buildDayBubbleContent(d, events, item.title, 'list-day');
      bubbles.push(bubble);
    });

    // always send carousel (will show "ไม่มีงาน" for empty days)
    const flex = {
      type: 'flex',
      altText: 'ตารางเช้า: วันนี้ / พรุ่งนี้ / มะรืน',
      contents: {
        type: 'carousel',
        contents: bubbles
      }
    };

    pushFlexToLine(flex);
  } catch (err) {
    console.error('sendMorningReminder error:', err);
  }
}

function createTimeTriggers() {
  const all = ScriptApp.getProjectTriggers();
  all.forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'sendAfternoonReminder' || fn === 'sendTomorrowReminder' || fn === 'sendMorningReminder') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // สร้าง trigger 11:00 -> sendAfternoonReminder
  ScriptApp.newTrigger('sendAfternoonReminder')
    .timeBased()
    .everyDays(1)
    .atHour(11)
    .create();

  // สร้าง trigger 17:00 -> sendTomorrowReminder
  ScriptApp.newTrigger('sendTomorrowReminder')
    .timeBased()
    .everyDays(1)
    .atHour(17)
    .create();

  // สร้าง trigger 6:00 -> sendMorningReminder
  ScriptApp.newTrigger('sendMorningReminder')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();

  Logger.log('Triggers created: sendMorningReminder@06:00, sendAfternoonReminder@11:00, sendTomorrowReminder@17:00');
}

// === FLEX BUILDERS & COLORS ===
function headerColorForAction(action) {
  switch ((action || '').toLowerCase()) {
    case 'create': return '#28a745'; // เพิ่ม = เขียวสด
    case 'delete': return '#dc3545'; // ลบ = แดงสด
    case 'update': return '#fd7e14'; // แก้ไข = ส้มสด
    case 'list':   return '#007bff'; // ดูตาราง = น้ำเงินสด (ทั่วไป)
    case 'info':   return '#1DB446'; // ข้อมูลทั่วไป
    default:       return '#1DB446';
  }
}

function getDayColorByWeekday(weekday) {
  switch (weekday) {
    case 0: return "#FF3B30"; // Sun
    case 1: return "#FFD60A"; // Mon
    case 2: return "#FF9ECF"; // Tue
    case 3: return "#34C759"; // Wed
    case 4: return "#FF9F0A"; // Thu
    case 5: return "#0A84FF"; // Fri
    case 6: return "#8E44AD"; // Sat
    default: return "#1DB446";
  }
}

// makeActionBubble now accepts optional dateObj to make bubble clickable to day view
function makeActionBubble(action, title, body, dateObj) {
  const color = headerColorForAction(action);
  const bubble = {
    type: "flex",
    altText: title,
    contents: {
      type: "bubble",
      header: {
        type: "box",
        layout: "vertical",
        contents: [{ type: "text", text: title, weight: "bold", size: "lg", color: "#ffffff" }],
        backgroundColor: color,
        paddingAll: "md"
      },
      body: {
        type: "box",
        layout: "vertical",
        spacing: "sm",
        contents: [{ type: "text", text: body, size: "sm", wrap: true }]
      }
    }
  };

  if (dateObj && dateObj instanceof Date && !isNaN(dateObj.getTime())) {
    const Y = dateObj.getFullYear();
    const M = ('0' + (dateObj.getMonth() + 1)).slice(-2);
    const D = ('0' + dateObj.getDate()).slice(-2);
    const dayUrl = `https://calendar.google.com/calendar/r/day/${Y}/${M}/${D}`;
    bubble.contents.action = { type: "uri", uri: dayUrl };
  }

  return bubble;
}

function makeCreateConfirmForEvent(ev, title, description, location) {
  const s = ev.getStartTime();
  const isAll = ev.isAllDayEvent && ev.isAllDayEvent();
  const timeLine = isAll ? `\nเวลา: ทั้งวัน` : `\nเวลา: ${formatTime(s)} - ${formatTime(ev.getEndTime())}`;
  const body = `เรื่อง: ${title}\nวันที่: ${formatDate(s)}${timeLine}${location ? `\nสถานที่: ${location}` : ''}${description ? `\nรายละเอียด: ${description}` : ''}`;
  return makeActionBubble('create', 'สร้างงานแล้ว', body, s);
}

function buildDayBubbleContent(dateObj, events, titleOverride, actionType) {
  const dateStr = formatDate(dateObj);
  const title = titleOverride || `ตารางวันที่ ${dateStr}`;
  const headerColor = (actionType === 'list-day') ? getDayColorByWeekday(dateObj.getDay()) : headerColorForAction(actionType || 'list');

  const rows = [];
  if (!events || events.length === 0) {
    rows.push({ type: "box", layout: "baseline", contents: [{ type: "text", text: "✅ ไม่มีงาน", size: "sm", color: "#999999" }] });
  } else {
    events.forEach(ev => {
      const isAllDay = ev.isAllDayEvent && ev.isAllDayEvent();
      const timeText = isAllDay ? "ทั้งวัน" : `${formatTime(ev.getStartTime())} - ${formatTime(ev.getEndTime())}`;

      rows.push({
        type: "box",
        layout: "horizontal",
        spacing: "md",
        contents: [
          {
            type: "box",
            layout: "vertical",
            contents: [{ type: "text", text: timeText, size: "sm", color: "#555555", wrap: false }],
            width: "36%" // เพิ่มพื้นที่ให้เวลามากขึ้น ป้องกันการถูกตัด
          },
          {
            type: "box",
            layout: "vertical",
            contents: [
              { type: "text", text: ev.getTitle(), size: "sm", color: "#111111", wrap: true },
              (ev.getLocation() ? { type: "text", text: `📍 ${ev.getLocation()}`, size: "xs", color: "#777777", wrap: true } : null)
            ].filter(Boolean),
            width: "64%"
          }
        ]
      });
    });
  }

  const Y = dateObj.getFullYear();
  const M = ('0' + (dateObj.getMonth() + 1)).slice(-2);
  const D = ('0' + dateObj.getDate()).slice(-2);
  const dayUrl = `https://calendar.google.com/calendar/r/day/${Y}/${M}/${D}`;

  return {
    type: "bubble",
    size: "mega",
    action: { type: "uri", uri: dayUrl },
    header: {
      type: "box",
      layout: "vertical",
      contents: [
        { type: "text", text: title, weight: "bold", size: "lg", color: "#ffffff" },
        { type: "text", text: dateStr, size: "sm", color: "#ffffffcc" }
      ],
      backgroundColor: headerColor,
      paddingAll: "md"
    },
    body: { type: "box", layout: "vertical", spacing: "sm", contents: rows }
  };
}

// === LINE helpers ===
function pushFlexToLine(flexMessage) {
  const url = 'https://api.line.me/v2/bot/message/push';
  const payload = { to: LINE_USER_ID, messages: [flexMessage] };
  UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + LINE_CHANNEL_ACCESS_TOKEN },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
}

// === Shortcuts Flex (for Rich Menu or on-demand) ===
function makeShortcutsFlex() {
  // colors: frame green, buttons: blue(today), yellow(tomorrow), red(next week)
  const frameColor = "#1DB446"; // green
  const btnTodayColor = "#4285F4"; // blue
  const btnTomorrowColor = "#FBBC05"; // yellow
  const btnNextWeekColor = "#EA4335"; // red

  const bubble = {
    type: "flex",
    altText: "คำสั่งลัด",
    contents: {
      type: "bubble",
      header: {
        type: "box",
        layout: "vertical",
        contents: [{ type: "text", text: "คำสั่งลัด", weight: "bold", size: "lg", color: "#ffffff" }],
        backgroundColor: frameColor,
        paddingAll: "md"
      },
      body: {
        type: "box",
        layout: "vertical",
        spacing: "sm",
        contents: [
          { type: "text", text: "กดเลือกคำสั่งลัดที่ต้องการ", size: "sm", wrap: true },
          {
            type: "box",
            layout: "vertical",
            spacing: "sm",
            contents: [
              {
                type: "button",
                style: "primary",
                color: btnTodayColor,
                action: { type: "message", label: "วันนี้มีอะไรบ้าง", text: "วันนี้มีอะไรบ้าง" }
              },
              {
                type: "button",
                style: "primary",
                color: btnTomorrowColor,
                action: { type: "message", label: "วันพรุ่งนี้ มีอะไรบ้าง", text: "วันพรุ่งนี้ มีอะไรบ้าง" }
              },
              {
                type: "button",
                style: "primary",
                color: btnNextWeekColor,
                action: { type: "message", label: "สัปดาห์หน้ามีอะไรบ้าง", text: "สัปดาห์หน้ามีอะไรบ้าง" }
              }
            ]
          }
        ]
      }
    }
  };

  return bubble;
}

// === Quick Replies helper (kept for reference) ===
function makeDefaultQuickReplies() {
  // Not used anymore by default; kept for backward compatibility if needed
  const buttons = [
    { label: "วันนี้มีอะไรบ้าง", text: "วันนี้มีอะไรบ้าง" },
    { label: "วันพรุ่งนี้ มีอะไรบ้าง", text: "วันพรุ่งนี้ มีอะไรบ้าง" },
    { label: "สัปดาห์หน้ามีอะไรบ้าง", text: "สัปดาห์หน้ามีอะไรบ้าง" }
  ];
  return buttons.map(b => ({ type: "action", action: { type: "message", label: b.label, text: b.text } }));
}

// === DATE/TIME formatting & helpers ===
function formatTime(date) { return Utilities.formatDate(date, "Asia/Bangkok", "HH:mm"); }
function formatDate(date) { return Utilities.formatDate(date, "Asia/Bangkok", "d MMMM yyyy"); }

// === Gemini AI integration ===
function callGemini(text) {
  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-lite:generateContent?key=${GEMINI_API_KEY}`;

  const prompt = `
คุณคือ "Thai Calendar Command Parser" ทำหน้าที่สกัดคำสั่งภาษาพูด/ภาษาธรรมชาติของผู้ใช้เป็น JSON (ห้ามใส่โค้ดบล็อก)
วันเวลาปัจจุบัน (โซน Asia/Bangkok): ${new Date().toISOString()}

ให้ตอบกลับเป็น JSON เท่านั้นตามสคีมา:
{
 "action": "create" | "delete" | "list" | "update" | "unknown",
 "title": string?,
 "startTime": "YYYY-MM-DD" or "YYYY-MM-DDTHH:MM:SS" or null,
 "endTime": "YYYY-MM-DDTHH:MM:SS" or null,
 "description": string?,
 "location": string?,
 "dayOfMonth": number?,
 "isRecurring": true|false,
 "newTitle": string?,
 "newStartTime": "YYYY-MM-DD" or "YYYY-MM-DDTHH:MM:SS" or null,
 "newEndTime": "YYYY-MM-DDTHH:MM:SS" or null
}
ข้อกำชับ: 
- หากผู้ใช้พูดว่า "ทุกวันที่ X" หรือ "ทุกเดือน" ให้ isRecurring=true. 
- หากพูดเป็น "วันที่ 27 เดือนนี้" ให้คืน startTime เป็นวันที่ 27 ของเดือนปัจจุบันในรูปแบบ "YYYY-MM-DD".
- หากพูดเป็น "23 สิงหา" ให้คืน startTime ในรูปแบบ "YYYY-MM-DD".
- สำหรับ action="update" ให้ใช้ "title" เป็นชื่อเดิมสำหรับค้นหา, "newTitle", "newStartTime", "newEndTime" สำหรับค่าที่จะอัพเดต, "description" และ "location" เป็นค่าใหม่ถ้ามี.
- ห้ามคืนข้อความธรรมดา หรือโค้ดอื่น ๆ — คืนเฉพาะ JSON เท่านั้น.

ข้อความผู้ใช้:
"${text}"
`;

  const payload = { contents: [{ parts: [{ text: prompt }] }], safetySettings: [
    { category: "HARM_CATEGORY_HARASSMENT", threshold: "BLOCK_NONE" },
    { category: "HARM_CATEGORY_HATE_SPEECH", threshold: "BLOCK_NONE" },
    { category: "HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold: "BLOCK_NONE" },
    { category: "HARM_CATEGORY_DANGEROUS_CONTENT", threshold: "BLOCK_NONE" }
  ] };

  const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };

  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const code = response.getResponseCode();
    const text = response.getContentText();
    if (code === 200) {
      const result = JSON.parse(text);
      if (result.candidates && result.candidates[0].content.parts[0].text) {
        const inner = result.candidates[0].content.parts[0].text.replace(/```json\n?/i, '').replace(/\n?```/g, '');
        try {
          return JSON.parse(inner);
        } catch (e) {
          console.error('Failed to parse Gemini inner JSON:', inner);
          return null;
        }
      }
    } else {
      console.error(`Gemini Error ${code}: ${text}`);
    }
    return null;
  } catch (err) {
    console.error('Gemini fetch error:', err);
    return null;
  }
}

// === Local fallback parser / heuristics ===
function tryLocalThaiParse(text) {
  const t = text.toLowerCase();
  if (/(มีอะไรบ้าง|ตาราง|คิว|ว่างไหม|free|available)/.test(t)) {
    const now = new Date();
    let d = new Date(now);
    if (/พรุ่งนี้/.test(t)) d.setDate(now.getDate() + 1);
    else if (/มะรืน/.test(t)) d.setDate(now.getDate() + 2);
    else if (/สัปดาห์หน้า/.test(t)) return { action: 'list', startTime: '[สัปดาห์หน้า]T00:00:00' };
    else if (/วันนี้/.test(t)) d = now;
    d.setHours(0,0,0,0);
    return { action: 'list', startTime: d.toISOString() };
  }

  // recurring monthly
  const rec = detectRecurringInText(text);
  if (rec && rec.type === 'monthly' && rec.day) {
    return { action: 'create', title: rec.title || ('นัดประจำเดือนวันที่ ' + rec.day), dayOfMonth: rec.day };
  }

  // quick create heuristics
  if (/เพิ่มนัด|ตั้งนัด|จอง/.test(t)) {
    const timeRange = t.match(/(\d{1,2}[:.]\d{2})\s*[-–]\s*(\d{1,2}[:.]\d{2})/);
    const dateLike = t.match(/(\d{1,2}\s*[ก-ฮ]+(?:\s*\d{2,4})?|\d{1,2}\/\d{1,2}\/?\d{0,4}|พรุ่งนี้|วันนี้|มะรืน)/);
    let titleMatch = t.replace(/เพิ่มนัด|ตั้งนัด|จอง/, '').trim();
    if (dateLike) titleMatch = titleMatch.replace(dateLike[0], '').trim();
    const info = { action: 'create', title: titleMatch || 'นัดใหม่' };
    if (timeRange && dateLike) {
      const s = parseTimeToToday(timeRange[1], dateLike[0]);
      const e = parseTimeToToday(timeRange[2], dateLike[0]);
      if (s && e) { info._parsedStart = s; info._parsedEnd = e; info._timeSpecified = true; }
    } else if (dateLike) {
      const d = parseDateFromThaiText(dateLike[0]);
      if (d) { info._parsedStart = d; info._timeSpecified = false; }
    }
    return info;
  }

  return { action: 'unknown' };
}

function detectRecurringInText(text) {
  const t = text.toLowerCase();
  const m1 = t.match(/ทุก\s*วันที่?\s*(\d{1,2})/);
  const m2 = t.match(/วันที่?\s*(\d{1,2})\s*ของทุก\s*เดือน/);
  if (m1) return { type: 'monthly', day: parseInt(m1[1],10) };
  if (m2) return { type: 'monthly', day: parseInt(m2[1],10) };
  if (/ทุก\s*เดือน/.test(t)) {
    const m = t.match(/(\d{1,2})/);
    if (m) return { type: 'monthly', day: parseInt(m[1],10) };
  }
  return null;
}

// === Date token resolution ===
function resolveParsedDateString(s, info) {
  if (!s || typeof s !== 'string') return null;
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/.test(s) || /^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  }
  const m = s.match(/^\[(.+?)\](?:T(\d{2}:\d{2}:\d{2}))?$/);
  if (m) {
    const token = m[1];
    const timePart = m[2];
    const base = interpretThaiRelativeDate(token);
    if (!base) return null;
    if (timePart) {
      const parts = timePart.split(':').map(p => parseInt(p,10));
      base.setHours(parts[0] || 0, parts[1] || 0, parts[2] || 0, 0);
    } else {
      base.setHours(0,0,0,0);
    }
    return base;
  }
  if (/พรุ่งนี้|วันนี้|มะรืน|เมื่อวาน|สัปดาห์หน้า|จันทร์หน้า|อังคารหน้า|พุธหน้า|พฤหัสหน้า|ศุกร์หน้า|เสาร์หน้า|อาทิตย์หน้า/.test(s)) {
    return interpretThaiRelativeDate(s);
  }
  if (info && info.action === 'list' && /^\d{1,2}$/.test(s)) {
    // สำหรับ list ถ้า startTime เป็นเลขวัน ให้ assume เดือนนี้
    const now = new Date();
    const day = parseInt(s, 10);
    const candidate = new Date(now.getFullYear(), now.getMonth(), day, 0,0,0);
    if (candidate.getDate() === day) return candidate;
  }
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function interpretThaiRelativeDate(token) {
  const now = new Date();
  const t = token.toLowerCase();
  const today = new Date(now); today.setHours(0,0,0,0);
  if (t.indexOf('วันนี้') !== -1) return today;
  if (t.indexOf('พรุ่งนี้') !== -1) { const d = new Date(today); d.setDate(d.getDate() + 1); return d; }
  if (t.indexOf('มะรืน') !== -1) { const d = new Date(today); d.setDate(d.getDate() + 2); return d; }
  if (t.indexOf('เมื่อวาน') !== -1) { const d = new Date(today); d.setDate(d.getDate() - 1); return d; }
  if (t.indexOf('สัปดาห์หน้า') !== -1) return getNextWeekMonday(today);
  const weekdays = { 'จันทร์':1, 'อังคาร':2, 'พุธ':3, 'พฤหัส':4, 'ศุกร์':5, 'เสาร์':6, 'อาทิตย์':0, 'อา':0 };
  for (const k in weekdays) {
    if (t.indexOf(k + 'หน้า') !== -1 || t.indexOf('วัน' + k + 'หน้า') !== -1) {
      return nextWeekdayDate(today, weekdays[k]);
    }
  }
  return null;
}

function getNextWeekMonday(date) {
  const d = new Date(date);
  const day = d.getDay();
  const daysUntilNextMonday = ((8 - day) % 7) || 7;
  d.setDate(d.getDate() + daysUntilNextMonday);
  d.setHours(0,0,0,0);
  return d;
}

function nextWeekdayDate(fromDate, targetWeekday) {
  const d = new Date(fromDate);
  const todayWd = d.getDay();
  let daysAhead = (targetWeekday - todayWd + 7) % 7;
  if (daysAhead === 0) daysAhead = 7;
  d.setDate(d.getDate() + daysAhead);
  d.setHours(0,0,0,0);
  return d;
}

// === Utility: time detection & simple thai date parsing ===
function detectTimeMention(text) {
  if (!text) return false;
  if (/\d{1,2}[:.]\d{2}/.test(text)) return true;
  if (/(\d{1,2}\s*โมง|บ่าย|เช้า|เย็น|ค่ำ|ดึก|ทุ่ม|เที่ยง|น\.)/i.test(text)) return true;
  return false;
}

function parseDateFromThaiText(text) {
  if (!text) return null;
  const t = text.toLowerCase();
  const thaiMonths = {
  'มกราคม':0,'ม.ค.':0,'มกรา':0,
  'กุมภาพันธ์':1,'ก.พ.':1,'กุมพา':1,'กุมภา':1,
  'มีนาคม':2,'มี.ค.':2,'มีนา':2,
  'เมษายน':3,'เม.ย.':3,'เมษา':3,
  'พฤษภาคม':4,'พ.ค.':4,'พฤษภา':4,
  'มิถุนายน':5,'มิ.ย.':5,'มิถุนา':5,
  'กรกฎาคม':6,'ก.ค.':6,'กรกฎา':6,
  'สิงหาคม':7,'ส.ค.':7,'สิงหา':7,
  'กันยายน':8,'ก.ย.':8,'กันยา':8,
  'ตุลาคม':9,'ต.ค.':9,'ตุลา':9,
  'พฤศจิกายน':10,'พ.ย.':10,'พฤศจิกา':10,
  'ธันวาคม':11,'ธ.ค.':11,'ธันวา':11,
  };

  // เพิ่มสำหรับ "วันที่ dd เดือนนี้"
  let m = t.match(/วันที่?\s*(\d{1,2})\s*เดือนนี้/);
  if (m) {
    const day = parseInt(m[1],10);
    const now = new Date();
    const month = now.getMonth();
    const year = now.getFullYear();
    const d = new Date(year, month, day, 0,0,0);
    if (d.getDate() === day) return d;
  }

  // dd/mm[/yyyy]
  m = t.match(/(\d{1,2})\s*\/\s*(\d{1,2})(?:\s*\/\s*(\d{2,4}))?/);
  if (m) {
    const day = parseInt(m[1],10);
    const month = parseInt(m[2],10) - 1;
    let year = new Date().getFullYear();
    if (m[3]) {
      year = parseInt(m[3],10);
      if (year < 100) { year += 2500; } // assume BE for 2-digit
      if (year > 2400) year -= 543; // BE -> AD
    }
    const d = new Date(year, month, day, 0,0,0);
    if (d.getDate() === day) return d;
  }

  // dd <thai month>
  m = t.match(/(\d{1,2})\s*(มกราคม|ม.ค\.?|กุมภาพันธ์|ก.พ\.?|มีนาคม|มี.ค\.?|เมษายน|เม.ย\.?|พฤษภาคม|พ.ค\.?|มิถุนายน|มิ.ย\.?|กรกฎาคม|ก.ค\.?|สิงหาคม|ส.ค\.?|กันยายน|ก.ย\.?|ตุลาคม|ต.ค\.?|พฤศจิกายน|พ.ย\.?|ธันวาคม|ธ.ค\.?)/);
  if (m) {
    const day = parseInt(m[1],10);
    const monthToken = m[2].replace(/\.$/, ''); // remove dot if any
    let monthIndex = thaiMonths[monthToken];
    const year = new Date().getFullYear();
    if (monthIndex !== undefined) {
      const d = new Date(year, monthIndex, day, 0,0,0);
      if (d.getDate() === day) return d;
    }
  }
  return null;
}

function parseTimeToToday(timeStr, dateToken) {
  const parts = timeStr.split(':');
  if (parts.length < 2) return null;
  const hh = parseInt(parts[0],10), mm = parseInt(parts[1],10);
  let base = new Date();
  const d = parseDateFromThaiText(dateToken);
  if (d) base = d;
  base.setHours(hh, mm, 0, 0);
  return base;
}

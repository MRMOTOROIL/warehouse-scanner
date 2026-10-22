/* 倉庫掃碼：操作人員／本機進度保護／可選日期歷史（2026-10-22 v2） */
(function () {
  'use strict';

  const STATE_KEY = 'warehouse-scanner:active-batch:v2';
  const OPERATOR_KEY = 'warehouse-scanner:last-operator:v1';
  let operatorName = '';
  let operators = [];
  let operatorDialog;
  let historyDialog;
  let selectButton;
  let currentSnapshot = null;
  let isRestoring = false;
  let selectedHistoryDate = taipeiDateToday();
  let historyRequestId = 0;

  function taipeiDateToday() {
    const parts = new Intl.DateTimeFormat('en-US', {
      timeZone: 'Asia/Taipei', year: 'numeric', month: '2-digit', day: '2-digit'
    }).formatToParts(new Date());
    const get = type => parts.find(p => p.type === type)?.value || '';
    return `${get('year')}-${get('month')}-${get('day')}`;
  }

  const safe = value => String(value ?? '').replace(/[&<>"']/g, character => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  }[character]));

  function getQtyScanned() {
    return Object.values(scannedItems).reduce((n, qty) => n + Number(qty || 0), 0);
  }

  function incompleteItems() {
    return Object.keys(requiredItems).map(name => {
      const required = Number(requiredItems[name]?.requiredQty || 0);
      const scanned = Number(scannedItems[name] || 0);
      return { name, required, scanned, missing: Math.max(0, required - scanned) };
    }).filter(item => item.missing > 0);
  }

  function readSnapshot() {
    try {
      const data = JSON.parse(localStorage.getItem(STATE_KEY) || 'null');
      if (!data || !Array.isArray(data.currentOrders)) return null;
      return data;
    } catch (_) { return null; }
  }

  function persistBatch() {
    if (isRestoring) return;
    try {
      if (!currentOrders.length || batchCompleted) {
        localStorage.removeItem(STATE_KEY);
        currentSnapshot = null;
        return;
      }
      currentSnapshot = {
        savedAt: new Date().toISOString(),
        operatorName,
        currentOrders: JSON.parse(JSON.stringify(currentOrders)),
        requiredItems: JSON.parse(JSON.stringify(requiredItems)),
        scannedItems: JSON.parse(JSON.stringify(scannedItems)),
        ignoredItems: JSON.parse(JSON.stringify(ignoredItems)),
        batchCompleted: false
      };
      localStorage.setItem(STATE_KEY, JSON.stringify(currentSnapshot));
    } catch (error) { console.warn('無法保存掃碼進度', error); }
  }

  function restoreBatch(snapshot) {
    if (!snapshot?.currentOrders?.length) return false;
    isRestoring = true;
    try {
      currentOrders = snapshot.currentOrders;
      requiredItems = snapshot.requiredItems || {};
      scannedItems = snapshot.scannedItems || {};
      ignoredItems = snapshot.ignoredItems || [];
      batchCompleted = false;
      renderAll();
      const pending = incompleteItems();
      const summary = pending.length
        ? pending.map(x => `${x.name} ${x.scanned}/${x.required}（少${x.missing}）`).join('；')
        : '未完成訂單已恢復，請核對商品';
      setStatus('waiting', '已恢復未完成批次', summary);
      return true;
    } finally {
      isRestoring = false;
    }
  }

  function makeDialog(id, title) {
    const overlay = document.createElement('div');
    overlay.id = id;
    overlay.setAttribute('role', 'dialog');
    overlay.setAttribute('aria-modal', 'true');
    overlay.style.cssText = 'display:none;position:fixed;inset:0;z-index:99999;background:rgba(0,0,0,.55);padding:20px;align-items:center;justify-content:center';
    const inner = document.createElement('section');
    inner.style.cssText = 'background:#fff;border-radius:18px;width:min(100%,620px);max-height:88vh;overflow:auto;padding:18px;color:#202124;box-shadow:0 5px 30px #0004';
    inner.innerHTML = `<h2 style="margin:0 0 14px;font-size:24px">${safe(title)}</h2><div class="modal-content"></div>`;
    overlay.appendChild(inner);
    document.body.appendChild(overlay);
    return { overlay, content:inner.querySelector('.modal-content') };
  }

  const show = dialog => { dialog.overlay.style.display = 'flex'; };
  const hide = dialog => { dialog.overlay.style.display = 'none'; };

  function actionButton(text, onClick, options = {}) {
    const button = document.createElement('button');
    button.type = 'button';
    button.textContent = text;
    button.style.cssText = 'padding:12px 14px;border:0;border-radius:12px;font-size:18px;font-weight:700;cursor:pointer;margin:5px;background:' + (options.background || '#1558a6') + ';color:' + (options.color || '#fff');
    button.disabled = Boolean(options.disabled);
    if (button.disabled) button.style.opacity = '.45';
    button.addEventListener('click', onClick);
    return button;
  }

  function showOperatorDialog() {
    const snapshot = readSnapshot();
    const pinnedName = snapshot?.currentOrders?.length && getTotalSnapshotScanned(snapshot) > 0 ? snapshot.operatorName : '';
    operatorDialog.content.replaceChildren();
    const help = document.createElement('p');
    help.style.cssText = 'line-height:1.5;font-size:16px';
    help.textContent = pinnedName
      ? `偵測到 ${pinnedName || '未指定'} 的未完成批次。為避免混淆操作紀錄，請先使用原操作人員完成。`
      : '請選擇這次操作的工作人員：';
    operatorDialog.content.appendChild(help);
    if (!operators.length) {
      operatorDialog.content.appendChild(document.createTextNode('無法取得操作人員清單。請確認 Google Sheets「操作人員」A 欄及 Apps Script 部署。'));
    }
    for (const name of operators) {
      operatorDialog.content.appendChild(actionButton(name, async () => {
        if (pinnedName && name !== pinnedName) return;
        operatorName = name;
        localStorage.setItem(OPERATOR_KEY, name);
        if (selectButton) selectButton.textContent = name;
        hide(operatorDialog);
        if (snapshot?.currentOrders?.length) {
          if (snapshot.operatorName && snapshot.operatorName !== name && getTotalSnapshotScanned(snapshot) > 0) return;
          const restored = restoreBatch(snapshot);
          persistBatch();
          if (restored) {
            try {
              await reloadBarcodeMap(null, true);
              const remaining = incompleteItems();
              setStatus('waiting', '已恢復未完成批次', remaining.map(x => `${x.name} ${x.scanned}/${x.required}（少${x.missing}）`).join('；') || '請核對商品');
            } catch (error) {
              showError('商品條碼載入失敗', '已恢復掃碼數量，但商品對照尚未載入，請檢查網路後重整');
            }
          }
        }
      }, { disabled: Boolean(pinnedName && name !== pinnedName) }));
    }
    const retry = actionButton('重新載入名單', async () => {
      await loadOperators();
      showOperatorDialog();
    }, { background:'#57616b' });
    operatorDialog.content.appendChild(retry);
    show(operatorDialog);
  }

  function getTotalSnapshotScanned(snapshot) {
    return Object.values(snapshot?.scannedItems || {}).reduce((n, v) => n + Number(v || 0), 0);
  }

  async function loadOperators() {
    try {
      const result = await apiCall('getWarehouseOperators');
      if (!result.ok) throw new Error(result.message || '取得名單失敗');
      operators = (result.operators || []).map(x => String(x).trim()).filter(Boolean);
    } catch (error) {
      console.warn('無法取得操作人員', error);
      operators = [];
    }
  }

  // 切換操作者不得接手已有商品掃描進度，避免操作歷程混淆。
  function switchOperator() {
    if (currentOrders.length && getQtyScanned() > 0 && !batchCompleted) {
      showError('不可切換操作人員', '目前批次已開始掃描商品，請由原操作人員完成');
      return;
    }
    showOperatorDialog();
  }

  function formatLocal(iso) {
    if (!iso) return '—';
    const date = new Date(iso);
    if (!Number.isFinite(date.getTime())) return String(iso);
    return date.toLocaleString('zh-TW', { timeZone:'Asia/Taipei', hour12:false, month:'2-digit', day:'2-digit', hour:'2-digit', minute:'2-digit', second:'2-digit' });
  }

  function makeHistoryToolbar() {
    const toolbar = document.createElement('div');
    toolbar.style.cssText = 'display:flex;flex-wrap:wrap;align-items:center;gap:8px;margin:5px 0 13px;';

    const fieldLabel = document.createElement('label');
    fieldLabel.htmlFor = 'warehouseHistoryDate';
    fieldLabel.textContent = '查詢日期';
    fieldLabel.style.cssText = 'font-size:16px;font-weight:700;';
    toolbar.appendChild(fieldLabel);

    const picker = document.createElement('input');
    picker.id = 'warehouseHistoryDate';
    picker.type = 'date';
    picker.value = selectedHistoryDate;
    picker.style.cssText = 'flex:1 1 160px;min-width:0;font-size:16px;padding:9px;border:1px solid #aab3bd;border-radius:10px;background:#fff;color:#202124;';
    picker.addEventListener('change', () => {
      if (/^\d{4}-\d{2}-\d{2}$/.test(picker.value)) {
        selectedHistoryDate = picker.value;
        loadHistory(selectedHistoryDate);
      }
    });
    toolbar.appendChild(picker);

    const today = actionButton('今天', () => {
      selectedHistoryDate = taipeiDateToday();
      picker.value = selectedHistoryDate;
      loadHistory(selectedHistoryDate);
    }, {background:'#57616b'});
    today.style.cssText += ';font-size:14px;padding:10px 12px;margin:0;';
    toolbar.appendChild(today);

    const refresh = actionButton('查詢', () => {
      if (!picker.value) return;
      selectedHistoryDate = picker.value;
      loadHistory(selectedHistoryDate);
    });
    refresh.style.cssText += ';font-size:14px;padding:10px 12px;margin:0;';
    toolbar.appendChild(refresh);
    return toolbar;
  }

  function renderHistoryRows(rows, date) {
    historyDialog.content.replaceChildren();
    historyDialog.content.appendChild(makeHistoryToolbar());
    const close = actionButton('返回掃碼', () => hide(historyDialog), {background:'#5c636e'});
    close.style.cssText += ';font-size:15px;padding:10px 12px;margin:0 0 10px;';
    historyDialog.content.appendChild(close);
    const title = document.createElement('p');
    title.style.cssText = 'font-size:14px;color:#667;margin:8px 0';
    title.textContent = `查詢日期：${date}。資料來源：Google Sheets「掃碼紀錄」；未完成不代表已出貨。`;
    historyDialog.content.appendChild(title);
    if (!rows.length) {
      const empty = document.createElement('p');
      empty.textContent = `${date} 沒有已記錄的掃碼資料。`;
      historyDialog.content.appendChild(empty);
      return;
    }
    for (const order of rows) {
      const card = document.createElement('div');
      card.style.cssText = 'padding:12px;border:1px solid #d9dee4;border-radius:12px;margin:12px 0;background:#fff';
      const header = document.createElement('div');
      header.innerHTML = `<strong style="font-size:18px;word-break:break-all">${safe(order.trackingNumber)}</strong><div style="margin-top:6px;font-size:14px">操作人員：${safe(order.operatorName || '未記錄')}｜${order.completedAt ? '✅ 完成' : '⏳ 未完成'}</div><div style="margin:6px 0;font-size:14px">完成時間：${safe(formatLocal(order.completedAt))}</div>`;
      card.appendChild(header);
      const table = document.createElement('table');
      table.style.cssText = 'width:100%;border-collapse:collapse;font-size:15px;table-layout:fixed';
      table.innerHTML = '<thead><tr><th style="text-align:left">應掃商品</th><th>數量</th><th>已掃</th></tr></thead><tbody>' +
        (order.items || []).map(item => `<tr><td style="border-top:1px solid #eee;padding:8px 2px;word-break:break-word">${safe(item.name)}</td><td style="text-align:center;border-top:1px solid #eee">${Number(item.scanned || 0)}/${Number(item.required || 0)}</td><td style="text-align:center;border-top:1px solid #eee;font-size:21px">${Number(item.required || 0) > 0 && Number(item.scanned || 0) >= Number(item.required || 0) ? '✅' : '⬜'}</td></tr>`).join('') + '</tbody>';
      card.appendChild(table);
      historyDialog.content.appendChild(card);
    }
  }

  async function loadHistory(date) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return;
    const requestId = ++historyRequestId;
    historyDialog.content.replaceChildren(makeHistoryToolbar());
    const loading = document.createElement('p');
    loading.textContent = `正在讀取 ${date} 的掃描紀錄…`;
    historyDialog.content.appendChild(loading);
    try {
      const result = await apiCall('getScanHistoryByDate', {date});
      if (requestId !== historyRequestId) return;
      if (!result.ok) throw new Error(result.message || '讀取歷史失敗');
      const rows = Array.isArray(result.orders) ? result.orders.slice() : [];
      // 本機當日未完成批次：僅加入今天的歷史，不污染其他日期。
      if (date === taipeiDateToday() && currentOrders.length && !batchCompleted) {
        const number = currentOrders.map(o => o.trackingNumber).join(',');
        const keyIndex = rows.findIndex(r => r.trackingNumber === number && r.operatorName === operatorName);
        const local = {
          trackingNumber: number, operatorName, completedAt: null,
          items: Object.keys(requiredItems).map(name => ({name, required: Number(requiredItems[name].requiredQty || 0), scanned: Number(scannedItems[name] || 0)}))
        };
        if (keyIndex < 0) rows.unshift(local);
        else if (!rows[keyIndex].completedAt) rows[keyIndex] = local;
      }
      renderHistoryRows(rows, date);
    } catch (error) {
      if (requestId !== historyRequestId) return;
      historyDialog.content.replaceChildren(makeHistoryToolbar());
      const el = document.createElement('p');
      el.style.color = '#b3261e';
      el.textContent = `讀取 ${date} 失敗：` + (error.message || String(error));
      historyDialog.content.appendChild(el);
      historyDialog.content.appendChild(actionButton('返回掃碼', () => hide(historyDialog), {background:'#57616b'}));
    }
  }

  async function openHistory() {
    selectedHistoryDate = taipeiDateToday();
    show(historyDialog);
    await loadHistory(selectedHistoryDate);
  }

  // 接管既有函式，但不改動託運單或商品本身的掃描規則。
  const oldClearBatch = clearBatch;
  clearBatch = function (showMessage = true) {
    if (showMessage) {
      if (!currentOrders.length) {
        setStatus('waiting', '目前沒有待清除的託運單', '請先掃描託運單');
        return false;
      }
      if (getQtyScanned() !== 0) {
        const detail = incompleteItems().map(x => `${x.name} 已掃 ${x.scanned}/${x.required}，還差 ${x.missing} 瓶`).join('；');
        showError('未掃滿，禁止清除批次', detail || '本批已有商品掃描紀錄，不可使用「清除」');
        return false;
      }
    } else if (!batchCompleted && currentOrders.length && getQtyScanned() > 0) {
      showError('未掃滿，禁止切換批次', '請先完成原批次');
      return false;
    }
    const result = oldClearBatch(showMessage);
    persistBatch();
    return result;
  };

  const oldAddTrackingResult = addTrackingResult;
  addTrackingResult = function (...args) {
    const result = oldAddTrackingResult(...args);
    persistBatch();
    return result;
  };

  const oldHandleKnownProductScan = handleKnownProductScan;
  handleKnownProductScan = function (...args) {
    const result = oldHandleKnownProductScan(...args);
    persistBatch();
    return result;
  };

  const oldHandleScan = handleScan;
  handleScan = function (...args) {
    if (!operatorName) {
      showOperatorDialog();
      return;
    }
    return oldHandleScan(...args);
  };

  // 覆蓋寫入請求：記錄 operatorName 到 Google Sheets I 欄。
  writeScanLog = function (barcode, shippingName, requiredQty, scannedQty) {
    const currentOperator = operatorName;
    const trackingNumbers = currentOrders.map(order => order.trackingNumber);
    const rawTexts = currentOrders.map(order => order.rawProductText);
    return apiCall('logScan', {
      data: {
        trackingNumber: trackingNumbers.join(','),
        rawProductText: rawTexts.join(' / '),
        shippingName,
        productCode: barcodeMap[barcode]?.productCode || '',
        requiredQty, scannedQty,
        status:'掃描成功',
        operatorName: currentOperator
      }
    }).then(result => {
      if (!result.ok) throw new Error(result.message || '掃碼紀錄寫入失敗');
    }).catch(error => {
      addHistory('', '⚠️ 掃碼紀錄寫入失敗：' + (error.message || String(error)), 'error');
    });
  };

  writeCompletionLog = async function () {
    const operatorForCompletion = operatorName;
    const trackingNumbers = currentOrders.map(order => order.trackingNumber);
    const rawProductTexts = currentOrders.map(order => order.rawProductText);
    const items = Object.keys(requiredItems).map(name => {
      const product = getProductByShippingName(name);
      return {
        shippingName:name,
        productCode: product?.productCode || '',
        requiredQty: requiredItems[name].requiredQty,
        scannedQty: Number(scannedItems[name] || 0)
      };
    });
    if (!items.length) return;
    try {
      const result = await apiCall('logBatchCompletion', {
        trackingNumbers, rawProductTexts, items,
        operatorName: operatorForCompletion
      });
      if (!result.ok) throw new Error(result.message || '完成紀錄寫入失敗');
    } catch(error) {
      addHistory('', '⚠️ 完成紀錄寫入失敗：' + (error.message || String(error)), 'error');
    }
  };

  // 底部四顆按鈕永遠同一列：清除｜歷史｜更新今日託運單｜操作人員。
  const bottomButtons = document.querySelector('.bottom-buttons');
  if (bottomButtons) {
    const buttons = [...bottomButtons.querySelectorAll('button')];
    const clearButton = buttons.find(b => b.getAttribute('onclick')?.includes('clearBatch'));
    const updateButton = buttons.find(b => b.getAttribute('onclick')?.includes('syncTodayData'));
    if (clearButton) clearButton.textContent = '清除';
    const historyButton = document.createElement('button');
    historyButton.className = 'btn btn-gray';
    historyButton.textContent = '歷史';
    historyButton.addEventListener('click', openHistory);
    selectButton = document.createElement('button');
    selectButton.className = 'btn btn-gray';
    selectButton.textContent = '選擇操作人員';
    selectButton.addEventListener('click', switchOperator);
    bottomButtons.replaceChildren(...[clearButton,historyButton,updateButton,selectButton].filter(Boolean));
    const style = document.createElement('style');
    style.textContent = `
      .bottom-bar .bottom-buttons {
        display: grid !important;
        grid-template-columns: minmax(0,.72fr) minmax(0,.72fr) minmax(0,1.65fr) minmax(0,1.02fr) !important;
        gap: 4px !important;
      }
      .bottom-bar .bottom-buttons > button {
        min-width: 0 !important;
        width: 100% !important;
        padding: 12px 1px !important;
        font-size: clamp(10px, 2.85vw, 14px) !important;
        line-height: 1.18 !important;
        white-space: nowrap !important;
        overflow: hidden;
        text-overflow: ellipsis;
      }
    `;
    document.head.appendChild(style);
    bottomButtons.style.gridTemplateColumns = 'minmax(0,.72fr) minmax(0,.72fr) minmax(0,1.65fr) minmax(0,1.02fr)';
    document.body.style.paddingBottom = '90px';
  }

  operatorDialog = makeDialog('warehouseOperatorDialog', '選擇操作人員');
  historyDialog = makeDialog('warehouseHistoryDialog', '掃描歷史');

  // 不依賴使用者必須先掃商品，託運單載入與任何商品成功掃描都會同步保存。
  window.addEventListener('pagehide', persistBatch);
  window.addEventListener('beforeunload', persistBatch);

  (async function start() {
    currentSnapshot = readSnapshot();
    await loadOperators();
    showOperatorDialog();
  })();
})();

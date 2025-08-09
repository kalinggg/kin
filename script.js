// 完整的前端API調用封裝
class QuotationAPI {
  constructor() {
    this.baseUrl = 'https://script.google.com/macros/s/AKfycbzlajZNy_D0j3dWDtZNj26Z9B42CAilvW1ucryfqZK9VsXbDt4Omv1FKbYzPL2slRI8wg/exec';
    this.password = 'your_secure_password';
  }
  
  async request(action, data = {}) {
    const params = new URLSearchParams({
      action,
      password: this.password,
      ...data
    });
    
    try {
      const response = await fetch(`${this.baseUrl}?${params}`, {
        method: 'POST',
        redirect: 'follow'
      });
      
      if (!response.ok) throw new Error('Network response was not ok');
      return await response.json();
    } catch (error) {
      console.error('API request failed:', error);
      throw error;
    }
  }
  
  async create(quotation) {
    return this.request('create', {
      number: quotation.number,
      date: quotation.date,
      customer: quotation.customer,
      contact: quotation.contact,
      address: quotation.address,
      notes: quotation.notes,
      items: JSON.stringify(quotation.items),
      total: quotation.total
    });
  }
  
  async getAll() {
    return this.request('read');
  }
  
  async getById(id) {
    const result = await this.request('read', { id });
    return result[0] || null;
  }
  
  async delete(id) {
    return this.request('delete', { id });
  }
}
// 全局變量
const api = new QuotationAPI();
let database = [];
let itemCounter = 1;

// DOM 加載完成後初始化
document.addEventListener('DOMContentLoaded', function() {
  initApp();
});

function initApp() {
  // 設置默認日期和報價單編號
  const today = new Date().toISOString().split('T')[0];
  document.getElementById('date').value = today;
  document.getElementById('quotation-number').value = generateQuotationNumber();

  // 從本地存儲加載數據
  const savedData = localStorage.getItem('quotationDatabase');
  database = savedData ? JSON.parse(savedData) : [];
  
  // 渲染報價單列表
  renderQuotationList();

  // 綁定事件監聽器
  setupEventListeners();
}

// 生成報價單編號
function generateQuotationNumber() {
  const now = new Date();
  const dateStr = now.toISOString().slice(0, 10).replace(/-/g, '');
  const randomNum = Math.floor(1000 + Math.random() * 9000);
  return `Q-${dateStr}-${randomNum}`;
}

// 設置事件監聽器
function setupEventListeners() {
  document.getElementById('add-item').addEventListener('click', addNewItem);
  document.getElementById('remove-item').addEventListener('click', removeSelectedItems);
  document.getElementById('save-btn').addEventListener('click', saveToDatabase);
  document.getElementById('delete-btn').addEventListener('click', deleteSelectedQuotation);
  document.getElementById('export-txt').addEventListener('click', exportToTxt);
  document.getElementById('export-excel').addEventListener('click', exportToExcel);
  document.getElementById('print-btn').addEventListener('click', printQuotation);
}

// 渲染報價單列表
function renderQuotationList() {
  const listContainer = document.getElementById('quotation-list');
  
  if (database.length === 0) {
    listContainer.innerHTML = '<p>尚未儲存任何報價單</p>';
    return;
  }

  listContainer.innerHTML = database.map(quotation => `
    <div class="quotation-item" data-id="${quotation.id}">
      <span>${quotation.number} - ${quotation.customer}</span>
      <span>總計: ${quotation.total}</span>
    </div>
  `).join('');

  // 為每個項目添加點擊事件
  document.querySelectorAll('.quotation-item').forEach(item => {
    item.addEventListener('click', function() {
      // 移除所有選中狀態
      document.querySelectorAll('.quotation-item').forEach(i => {
        i.classList.remove('selected');
      });
      // 添加當前選中狀態
      this.classList.add('selected');
      
      const quotationId = this.dataset.id;
      const quotation = database.find(q => q.id === quotationId);
      if (quotation && confirm('載入此報價單？現有內容將會被替換。')) {
        loadQuotation(quotation);
      }
    });
  });
}

// 使用範例
const api = new QuotationAPI();

// 在前端添加重試機制
async function loadWithRetry(retries = 3) {
  try {
    return await loadFromDatabase();
  } catch (error) {
    if (retries > 0) {
      await new Promise(resolve => setTimeout(resolve, 1000));
      return loadWithRetry(retries - 1);
    }
    throw error;
  }
}

// 從Google Sheets加載報價單
async function loadFromDatabase() {
    try {
        const data = await api.getAll();
        
        // 轉換數據格式
        database = data.map(record => ({
            id: record.id,
            number: record.number,
            date: record.date,
            customer: record.customer,
            contact: record.contact,
            address: record.address,
            notes: record.notes,
            items: JSON.parse(record.items || '[]'),
            total: record.total,
            created: record.created
        }));
        
        renderQuotationList();
    } catch (error) {
        alert('加載失敗: ' + error.message);
        console.error('Error:', error);
    }
}

document.addEventListener('DOMContentLoaded', function() {
    // 自動生成報價單編號 (格式: Q-年月日-序號)
    function generateQuotationNumber() {
        const now = new Date();
        const dateStr = now.toISOString().slice(0, 10).replace(/-/g, '');
        const randomNum = Math.floor(1000 + Math.random() * 9000);
        return `Q-${dateStr}-${randomNum}`;
    }

    // 設置默認日期和報價單編號
    const today = new Date().toISOString().split('T')[0];
    document.getElementById('date').value = today;
    document.getElementById('quotation-number').value = generateQuotationNumber();

    // 初始化項目計數器和資料庫
    let itemCounter = 1;
    let database = JSON.parse(localStorage.getItem('quotationDatabase')) || [];
    renderQuotationList();

    // 事件監聽器
    document.getElementById('add-item').addEventListener('click', addNewItem);
    document.getElementById('remove-item').addEventListener('click', removeSelectedItems);
    document.getElementById('save-btn').addEventListener('click', saveToDatabase);
    document.getElementById('delete-btn').addEventListener('click', deleteSelectedQuotation);
    document.getElementById('export-txt').addEventListener('click', exportToTxt);
    document.getElementById('export-excel').addEventListener('click', exportToExcel);
    document.getElementById('print-btn').addEventListener('click', printQuotation);

    // 添加新項目
    function addNewItem() {
        const tbody = document.getElementById('items-body');
        const newRow = document.createElement('tr');
        newRow.dataset.itemId = itemCounter;
        
        newRow.innerHTML = `
            <td class="item-number">${itemCounter}</td>
            <td><input type="text" class="item-description" placeholder="請輸入項目描述"></td>
            <td><input type="number" class="item-quantity" min="1" value="1"></td>
            <td><input type="number" class="item-price" min="0" placeholder="0"></td>
            <td class="item-total">0</td>
            <td><input type="text" class="item-remark" placeholder="備註"></td>
            <td><input type="checkbox" class="item-selector"></td>
        `;
        
        tbody.appendChild(newRow);
        itemCounter++;
        
        // 事件監聽器
        const qtyInput = newRow.querySelector('.item-quantity');
        const priceInput = newRow.querySelector('.item-price');
        qtyInput.addEventListener('input', calculateRowTotal);
        priceInput.addEventListener('input', calculateRowTotal);
        
        newRow.addEventListener('click', function(e) {
            if (e.target.type !== 'checkbox' && e.target.tagName !== 'INPUT') {
                const checkbox = this.querySelector('.item-selector');
                if (checkbox) {
                    checkbox.checked = !checkbox.checked;
                    this.classList.toggle('selected', checkbox.checked);
                }
            }
        });
        
        newRow.querySelector('.item-description').focus();
    }

    // 計算單行總額
    function calculateRowTotal() {
        const row = this.closest('tr');
        const qty = parseInt(row.querySelector('.item-quantity').value) || 0;
        const price = parseInt(row.querySelector('.item-price').value) || 0;
        row.querySelector('.item-total').textContent = qty * price;
        calculateTotals();
    }

    // 計算總計
    function calculateTotals() {
        const totals = Array.from(document.querySelectorAll('.item-total')).map(el => parseInt(el.textContent));
        document.getElementById('grand-total').textContent = totals.reduce((sum, num) => sum + num, 0);
    }

    // 刪除選中項目
    function removeSelectedItems() {
        const selectedItems = document.querySelectorAll('#items-body tr.selected');
        if (selectedItems.length === 0) return alert('請先選擇要刪除的項目！');
        
        if (confirm(`確定要刪除選中的 ${selectedItems.length} 個項目嗎？`)) {
            selectedItems.forEach(item => item.remove());
            updateItemNumbers();
            calculateTotals();
        }
    }

    // 更新項目序號
    function updateItemNumbers() {
        document.querySelectorAll('#items-body tr').forEach((row, index) => {
            row.querySelector('.item-number').textContent = index + 1;
        });
    }

    // 儲存到資料庫
async function saveToDatabase() {
    showLoading();
    try {
        const quotation = getCurrentQuotation();
        const result = await api.create(quotation);
        
        if (result && result.success) {
            // 重新從伺服器加載最新資料
            await loadFromDatabase();
            showSuccess('報價單已成功保存！');
            
            // 同時更新本地存儲
            database.push(quotation);
            localStorage.setItem('quotationDatabase', JSON.stringify(database));
            
            // 重置表單
            resetForm();
        } else {
            throw new Error(result?.message || '保存失敗');
        }
    } catch (error) {
        showError('保存失敗: ' + error.message);
        console.error('保存錯誤:', error);
    } finally {
        hideLoading();
    }
}

// 從資料庫加載報價單
async function loadFromDatabase() {
    showLoading();
    try {
        const data = await api.getAll();
        
        // 更新本地資料庫
        database = data.map(record => ({
            id: record.id,
            number: record.number,
            date: record.date,
            customer: record.customer,
            contact: record.contact,
            address: record.address,
            notes: record.notes,
            items: JSON.parse(record.items || '[]'),
            total: record.total,
            created: record.created
        }));
        
        // 更新本地存儲
        localStorage.setItem('quotationDatabase', JSON.stringify(database));
        
        // 重新渲染列表
        renderQuotationList();
    } catch (error) {
        showError('加載失敗: ' + error.message);
        console.error('加載錯誤:', error);
        
        // 嘗試從本地存儲恢復
        const localData = localStorage.getItem('quotationDatabase');
        if (localData) {
            database = JSON.parse(localData);
            renderQuotationList();
        }
    } finally {
        hideLoading();
    }
}

// 重置表單
function resetForm() {
    document.getElementById('items-body').innerHTML = '';
    itemCounter = 1;
    document.getElementById('quotation-number').value = generateQuotationNumber();
    document.getElementById('date').value = new Date().toISOString().split('T')[0];
    document.getElementById('customer').value = '';
    document.getElementById('contact-person').value = '';
    document.getElementById('address').value = '';
    document.getElementById('notes').value = '';
    document.getElementById('grand-total').textContent = '0';
}

// 渲染報價單列表 (改進版)
function renderQuotationList() {
    const listContainer = document.getElementById('quotation-list');
    
    if (!database || database.length === 0) {
        listContainer.innerHTML = '<p>尚未儲存任何報價單</p>';
        return;
    }

    // 按創建時間降序排序
    const sortedData = [...database].sort((a, b) => {
        return new Date(b.created || b.date) - new Date(a.created || a.date);
    });

    listContainer.innerHTML = sortedData.map(quotation => `
        <div class="quotation-item" data-id="${quotation.id}">
            <span>${quotation.number} - ${quotation.customer}</span>
            <span>總計: ${quotation.total}</span>
            <small>${formatDate(quotation.date)}</small>
        </div>
    `).join('');

    // 綁定點擊事件
    document.querySelectorAll('.quotation-item').forEach(item => {
        item.addEventListener('click', function() {
            document.querySelectorAll('.quotation-item').forEach(i => {
                i.classList.remove('selected');
            });
            this.classList.add('selected');
        });
    });
}

// 格式化日期
function formatDate(dateString) {
    if (!dateString) return '';
    const date = new Date(dateString);
    return date.toLocaleDateString('zh-TW');
}

    // 刪除選中報價單
    async function deleteSelectedQuotation() {
        const selectedQuotation = document.querySelector('.quotation-item.selected');
        if (!selectedQuotation) return alert('請先選擇要刪除的報價單！');
        
        const quotationId = selectedQuotation.dataset.id;
        if (confirm('確定要刪除此報價單嗎？此操作無法復原！')) {
            try {
                await api.delete(quotationId);
                await loadFromDatabase();
                showSuccess('報價單已刪除！');
            } catch (error) {
                showError('刪除失敗: ' + error.message);
            }
        }
    }

    // 渲染報價單列表
    function renderQuotationList() {
        const listContainer = document.getElementById('quotation-list');
        listContainer.innerHTML = database.length === 0 
            ? '<p>尚未儲存任何報價單</p>' 
            : database.map((quotation, index) => `
                <div class="quotation-item" data-id="${quotation.id}">
                    <span>${quotation.number} - ${quotation.customer}</span>
                    <span>總計: ${quotation.total}</span>
                </div>
            `).join('');
        
        // 添加點擊事件
        document.querySelectorAll('.quotation-item').forEach(item => {
            item.addEventListener('click', function() {
                // 移除所有選中狀態
                document.querySelectorAll('.quotation-item').forEach(i => {
                    i.classList.remove('selected');
                });
                // 添加當前選中狀態
                this.classList.add('selected');
                
                const quotationId = this.dataset.id;
                const quotation = database.find(q => q.id === quotationId);
                if (quotation && confirm('載入此報價單？現有內容將會被替換。')) {
                    loadQuotation(quotation);
                }
            });
        });
    }

    // 載入報價單
    function loadQuotation(quotation) {
        document.getElementById('items-body').innerHTML = '';
        itemCounter = 1;
        
        // 填充基本信息
        document.getElementById('quotation-number').value = quotation.number;
        document.getElementById('date').value = quotation.date;
        document.getElementById('customer').value = quotation.customer;
        document.getElementById('contact-person').value = quotation.contact;
        document.getElementById('address').value = quotation.address;
        document.getElementById('notes').value = quotation.notes;
        
        // 添加項目
        quotation.items.forEach(item => {
            const tbody = document.getElementById('items-body');
            const newRow = document.createElement('tr');
            newRow.dataset.itemId = itemCounter;
            
            newRow.innerHTML = `
                <td class="item-number">${itemCounter}</td>
                <td><input type="text" class="item-description" value="${item.description}"></td>
                <td><input type="number" class="item-quantity" min="1" value="${item.quantity}"></td>
                <td><input type="number" class="item-price" min="0" value="${item.price}"></td>
                <td class="item-total">${item.total}</td>
                <td><input type="text" class="item-remark" value="${item.remark}"></td>
                <td><input type="checkbox" class="item-selector"></td>
            `;
            
            tbody.appendChild(newRow);
            itemCounter++;
            
            // 事件監聽器
            const qtyInput = newRow.querySelector('.item-quantity');
            const priceInput = newRow.querySelector('.item-price');
            qtyInput.addEventListener('input', calculateRowTotal);
            priceInput.addEventListener('input', calculateRowTotal);
            
            newRow.addEventListener('click', function(e) {
                if (e.target.type !== 'checkbox' && e.target.tagName !== 'INPUT') {
                    const checkbox = this.querySelector('.item-selector');
                    if (checkbox) {
                        checkbox.checked = !checkbox.checked;
                        this.classList.toggle('selected', checkbox.checked);
                    }
                }
            });
        });
        
        document.getElementById('grand-total').textContent = quotation.total;
    }

    // 獲取當前報價單數據
    function getCurrentQuotation() {
        const items = Array.from(document.querySelectorAll('#items-body tr')).map(row => ({
            description: row.querySelector('.item-description').value || '未填寫',
            quantity: row.querySelector('.item-quantity').value,
            price: row.querySelector('.item-price').value,
            total: row.querySelector('.item-total').textContent,
            remark: row.querySelector('.item-remark').value || ''
        }));
        
        return {
            id: Date.now().toString(),
            number: document.getElementById('quotation-number').value,
            date: document.getElementById('date').value,
            customer: document.getElementById('customer').value || '未指定',
            contact: document.getElementById('contact-person').value || '未指定',
            address: document.getElementById('address').value || '未提供',
            notes: document.getElementById('notes').value || '',
            items: items,
            total: document.getElementById('grand-total').textContent
        };
    }

    // 導出TXT
    function exportToTxt() {
        const quotation = getCurrentQuotation();
        let txtContent = `=== 報價單 ===\n\n`;
        txtContent += `報價單編號: ${quotation.number}\n`;
        txtContent += `日期: ${quotation.date}\n`;
        txtContent += `客戶名稱: ${quotation.customer}\n`;
        txtContent += `聯絡人: ${quotation.contact}\n`;
        txtContent += `地址: ${quotation.address}\n\n`;
        
        txtContent += `項目列表:\n`;
        txtContent += `----------------------------------------------\n`;
        txtContent += `序號 | 項目描述 | 數量 | 單價 | 金額 | 備註\n`;
        txtContent += `----------------------------------------------\n`;
        
        quotation.items.forEach((item, index) => {
            txtContent += `${index + 1} | ${item.description} | ${item.quantity} | ${item.price} | ${item.total} | ${item.remark}\n`;
        });
        
        txtContent += `\n總計: ${quotation.total}\n\n`;
        txtContent += `備註:\n${quotation.notes}\n`;
        
        const blob = new Blob([txtContent], { type: 'text/plain;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `報價單_${quotation.number}.txt`;
        a.click();
        URL.revokeObjectURL(url);
    }

    // 導出Excel
    function exportToExcel() {
        const quotation = getCurrentQuotation();
        const wb = XLSX.utils.book_new();
        
        const wsData = [
            ['報價單', '', '', '', ''],
            ['報價單編號:', quotation.number, '', '日期:', quotation.date],
            ['客戶名稱:', quotation.customer, '', '聯絡人:', quotation.contact],
            ['地址:', quotation.address, '', '', ''],
            ['', '', '', '', ''],
            ['序號', '項目描述', '數量', '單價', '金額', '備註']
        ];
        
        quotation.items.forEach((item, index) => {
            wsData.push([
                index + 1,
                item.description,
                item.quantity,
                item.price,
                item.total,
                item.remark
            ]);
        });
        
        wsData.push(['', '', '', '總計:', quotation.total, '']);
        wsData.push(['', '', '', '', '', '']);
        wsData.push(['備註:', quotation.notes, '', '', '', '']);
        
        const ws = XLSX.utils.aoa_to_sheet(wsData);
        XLSX.utils.book_append_sheet(wb, ws, "報價單");
        XLSX.writeFile(wb, `報價單_${quotation.number}.xlsx`);
    }

    // 列印報價單
    function printQuotation() {
        // 準備列印內容
        const printContent = document.createElement('div');
        printContent.innerHTML = `
            <div style="font-family: 'Microsoft JhengHei', Arial, sans-serif; max-width: 900px; margin: 0 auto;">
                <h2 style="text-align: center; color: #3498db; border-bottom: 2px solid #f5f5f5; padding-bottom: 15px; margin-bottom: 30px;">
                    報價單
                </h2>
                
                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 30px;">
                    <div>
                        <p><strong>報價單編號:</strong> ${document.getElementById('quotation-number').value}</p>
                        <p><strong>日期:</strong> ${document.getElementById('date').value}</p>
                    </div>
                    <div>
                        <p><strong>客戶名稱:</strong> ${document.getElementById('customer').value || '未指定'}</p>
                        <p><strong>聯絡人:</strong> ${document.getElementById('contact-person').value || '未指定'}</p>
                    </div>
                </div>
                
                <p><strong>地址:</strong> ${document.getElementById('address').value || '未提供'}</p>
                
                <table style="width: 100%; border-collapse: collapse; margin: 25px 0; border: 1px solid #ddd;">
                    <thead>
                        <tr style="background-color: #3498db; color: white; text-align: left;">
                            <th style="padding: 12px 15px; width: 5%;">序號</th>
                            <th style="padding: 12px 15px; width: 35%;">項目描述</th>
                            <th style="padding: 12px 15px; width: 10%;">數量</th>
                            <th style="padding: 12px 15px; width: 15%;">單價</th>
                            <th style="padding: 12px 15px; width: 15%;">金額</th>
                            <th style="padding: 12px 15px; width: 15%;">備註</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${Array.from(document.querySelectorAll('#items-body tr')).map(row => `
                            <tr style="border-bottom: 1px solid #ddd;">
                                <td style="padding: 12px 15px;">${row.querySelector('.item-number').textContent}</td>
                                <td style="padding: 12px 15px;">${row.querySelector('.item-description').value || '未填寫'}</td>
                                <td style="padding: 12px 15px; text-align: right;">${row.querySelector('.item-quantity').value}</td>
                                <td style="padding: 12px 15px; text-align: right;">${row.querySelector('.item-price').value || 0}</td>
                                <td style="padding: 12px 15px; text-align: right;">${row.querySelector('.item-total').textContent}</td>
                                <td style="padding: 12px 15px;">${row.querySelector('.item-remark').value || ''}</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
                
                <div style="text-align: right; margin-top: 20px; padding: 15px; background-color: #f5f5f5; border-radius: 4px;">
                    <div style="font-size: 18px; font-weight: bold;">總計: ${document.getElementById('grand-total').textContent}</div>
                </div>
                
                ${document.getElementById('notes').value ? `
                    <div style="margin-top: 30px;">
                        <h3 style="margin-bottom: 10px;">備註說明</h3>
                        <p style="white-space: pre-line;">${document.getElementById('notes').value}</p>
                    </div>
                ` : ''}
                
                <div style="margin-top: 50px; display: flex; justify-content: space-between;">
                    <div style="width: 200px; border-top: 1px solid #333; padding-top: 10px; text-align: center;">
                        客戶簽名
                    </div>
                    <div style="width: 200px; border-top: 1px solid #333; padding-top: 10px; text-align: center;">
                        公司簽章
                    </div>
                </div>
            </div>
        `;
        
        // 創建列印窗口
        const printWindow = window.open('', '_blank');
        printWindow.document.write(`
            <!DOCTYPE html>
            <html>
                <head>
                    <title>報價單 - ${document.getElementById('customer').value || '未命名客戶'}</title>
                    <meta charset="UTF-8">
                    <style>
                        body { margin: 0; padding: 20px; font-family: 'Microsoft JhengHei', Arial, sans-serif; }
                        @page { size: A4; margin: 10mm; }
                    </style>
                </head>
                <body>
                    ${printContent.innerHTML}
                    <script>
                        window.onload = function() {
                            setTimeout(function() {
                                window.print();
                                window.close();
                            }, 200);
                        };
                    <\/script>
                </body>
            </html>
        `);
        printWindow.document.close();
    }

    // 輔助函數 (需自行實現)
    function showLoading() {
        console.log('顯示載入中...');
        // 實際實現應顯示載入指示器
    }
    
    function hideLoading() {
        console.log('隱藏載入中...');
        // 實際實現應隱藏載入指示器
    }
    
    function showSuccess(message) {
        alert(message);
        // 實際實現應顯示成功通知
    }
    
    function showError(message) {
        alert(message);
        // 實際實現應顯示錯誤通知
    }
});
// 加載狀態函數
function showLoading() {
  document
.getElementById('loading').style.display = 'flex';
}

function hideLoading() {
  document
.getElementById('loading').style.display = 'none';
}

function showSuccess(message) {
  const errorDiv = document.getElementById('error-message');
  errorDiv
.style.display = 'block';
  errorDiv
.style.backgroundColor = '#d4edda';
  errorDiv
.style.color = '#155724';
  errorDiv
.style.borderColor = '#c3e6cb';
  errorDiv
.textContent = message;
  setTimeout(() => errorDiv.style.display = 'none', 3000);
}

function showError(message) {
  const errorDiv = document.getElementById('error-message');
  errorDiv
.style.display = 'block';
  errorDiv
.style.backgroundColor = '#f8d7da';
  errorDiv
.style.color = '#721c24';
  errorDiv
.style.borderColor = '#f5c6cb';
  errorDiv
.textContent = message;
  setTimeout(() => errorDiv.style.display = 'none', 5000);
}

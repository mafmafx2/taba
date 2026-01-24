const RANKING_URL = './data/pts-ranking.json';
const DISCLOSURE_URL = './data/disclosures.json';
const REFRESH_INTERVAL_MS = 60000;

const rankingBody = document.getElementById('ranking-body');
const status = document.getElementById('status');
const lastUpdated = document.getElementById('last-updated');
const refreshButton = document.getElementById('refresh-button');

const formatNumber = (value) => {
  if (value === null || value === undefined) return '-';
  return Number(value).toLocaleString('ja-JP');
};

const formatChange = (value) => {
  if (value === null || value === undefined) return '-';
  const sign = value > 0 ? '+' : '';
  return `${sign}${value.toFixed(2)}`;
};

const buildDisclosureMap = (disclosures) => {
  const map = new Map();
  disclosures
    .sort((a, b) => new Date(b.publishedAt) - new Date(a.publishedAt))
    .forEach((disclosure) => {
      if (!map.has(disclosure.code)) {
        map.set(disclosure.code, []);
      }
      if (map.get(disclosure.code).length < 3) {
        map.get(disclosure.code).push(disclosure);
      }
    });
  return map;
};

const renderRanking = (ranking, disclosureMap) => {
  rankingBody.innerHTML = '';
  ranking.forEach((item, index) => {
    const tr = document.createElement('tr');
    const changeClass = item.change > 0 ? 'positive' : item.change < 0 ? 'negative' : '';

    const disclosureList = disclosureMap.get(item.code) || [];
    const disclosureHtml = disclosureList.length
      ? `<ul class="disclosure-list">${disclosureList
          .map(
            (disclosure) =>
              `<li><a href="${disclosure.url}" target="_blank" rel="noreferrer">${disclosure.title}</a><span>${new Date(
                disclosure.publishedAt,
              ).toLocaleString('ja-JP')}</span></li>`,
          )
          .join('')}</ul>`
      : '<span class="muted">該当なし</span>';

    tr.innerHTML = `
      <td>${index + 1}</td>
      <td>${item.code}</td>
      <td>${item.name}</td>
      <td>¥${formatNumber(item.price)}</td>
      <td class="${changeClass}">${formatChange(item.change)}</td>
      <td>${formatNumber(item.volume)}</td>
      <td>${disclosureHtml}</td>
    `;

    rankingBody.appendChild(tr);
  });
};

const setStatus = (message, type = 'info') => {
  status.textContent = message;
  status.className = `status ${type}`;
};

const updateTimestamp = () => {
  lastUpdated.textContent = new Date().toLocaleString('ja-JP');
};

const fetchData = async () => {
  try {
    setStatus('更新中...', 'loading');
    const [rankingResponse, disclosureResponse] = await Promise.all([
      fetch(RANKING_URL),
      fetch(DISCLOSURE_URL),
    ]);

    if (!rankingResponse.ok || !disclosureResponse.ok) {
      throw new Error('データ取得に失敗しました');
    }

    const rankingData = await rankingResponse.json();
    const disclosureData = await disclosureResponse.json();
    const disclosureMap = buildDisclosureMap(disclosureData.items);

    renderRanking(rankingData.items, disclosureMap);
    updateTimestamp();
    setStatus('更新完了', 'success');
  } catch (error) {
    setStatus(`エラー: ${error.message}`, 'error');
  }
};

refreshButton.addEventListener('click', fetchData);

fetchData();
setInterval(fetchData, REFRESH_INTERVAL_MS);

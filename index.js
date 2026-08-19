/**
 * @OnlyCurrentDoc
 */

class MultiChainNftTraitViewer {
  static get CONFIG_SHEET_NAME() { return 'Config'; }
  static get CACHE_SHEET_NAME() { return 'Cache'; }

  static get CHAIN_SLUG_MAP() {
    return {
      'eth-mainnet': 'ethereum',
      'polygon-mainnet': 'matic',
      'matic-mainnet': 'matic',
      'arb-mainnet': 'arbitrum',
      'opt-mainnet': 'optimism',
      'base-mainnet': 'base',
      'avax-mainnet': 'avalanche',
      'bnb-mainnet': 'bsc'
    };
  }

  static setupConfigSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(MultiChainNftTraitViewer.CONFIG_SHEET_NAME) ?? ss.insertSheet(MultiChainNftTraitViewer.CONFIG_SHEET_NAME);
    sheet.clear();

    sheet.getRange('A1:B4').setValues([
      ['Network (EVM / XRPL):', 'XRPL'],
      ['Endpoint (Alchemy or XRPL Node):', 'https://s1.ripple.com:51234'],
      ['Owner Address:', ''],
      ['Contract / Issuer Address:', '']
    ]);
    sheet.getRange('A5').setValue('Traits to Display (one per cell below):');
    sheet.getRange('A1:A5').setFontWeight('bold');

    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['EVM', 'XRPL'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange('B1').setDataValidation(rule);

    SpreadsheetApp.getUi().alert([
      'Config sheet has been set up.',
      '1. Select Network in cell B1 (EVM or XRPL).',
      '2. Enter Endpoint in cell B2:',
      '   - EVM: https://eth-mainnet.g.alchemy.com/v2/  (use polygon-mainnet / arb-mainnet etc.)',
      '   - XRPL: https://s1.ripple.com:51234',
      '3. Enter Owner Address (B3), Contract/Issuer Address (B4), and Traits starting from cell A6.'
    ].join('\n'));
  }

  static fetchNftDataWithRefreshedCache() {
    new MultiChainNftTraitViewer({ ss: SpreadsheetApp.getActiveSpreadsheet(), ui: SpreadsheetApp.getUi(), useCache: false, refresh: true }).build();
  }

  static fetchNftDataWithCache() {
    new MultiChainNftTraitViewer({ ss: SpreadsheetApp.getActiveSpreadsheet(), ui: SpreadsheetApp.getUi(), useCache: true }).build();
  }

  static fetchNftDataWithoutCache() {
    new MultiChainNftTraitViewer({ ss: SpreadsheetApp.getActiveSpreadsheet(), ui: SpreadsheetApp.getUi(), useCache: false }).build();
  }

  static columnToLetter(column) {
    if (!Number.isInteger(column) || column < 1) {
      throw new Error('Column number must be a positive integer.');
    }
    let letter = '';
    let n = column;
    while (n > 0) {
      const remainder = (n - 1) % 26;
      letter = String.fromCharCode(65 + remainder) + letter;
      n = Math.floor((n - 1) / 26);
    }
    return letter;
  }

  static resolveOpenSeaChainSlug(endpoint) {
    const map = MultiChainNftTraitViewer.CHAIN_SLUG_MAP;
    const key = Object.keys(map).find(k => (endpoint || '').includes(k));
    return key ? map[key] : 'ethereum';
  }

  static hexToUtf8(hex) {
    if (!hex) return '';
    try {
      const bytes = [];
      for (let c = 0; c < hex.length; c += 2) {
        bytes.push(parseInt(hex.substr(c, 2), 16));
      }
      return Utilities.newBlob(bytes).getDataAsString('UTF-8');
    } catch (e) {
      return '';
    }
  }

  constructor({ ss, ui, useCache, refresh = false }) {
    Object.assign(this, { ss, ui, useCache, refresh });
    this.init();
  }

  init() {
    const configSheet = this.ss.getSheetByName(MultiChainNftTraitViewer.CONFIG_SHEET_NAME);
    if (!configSheet) throw new Error(`Sheet "${MultiChainNftTraitViewer.CONFIG_SHEET_NAME}" not found. Please run "Setup Config Sheet" first.`);

    [this.network, this.endpoint, this.ownerAddress, this.contractAddress] = configSheet.getRange('B1:B4').getValues().flat();
    this.network = String(this.network || 'EVM').toUpperCase();

    if (!this.ownerAddress) throw new Error('Owner Address must be entered in the Config sheet.');
    if (this.network === 'EVM' && !this.contractAddress) throw new Error('Contract Address must be entered in the Config sheet for EVM.');
    if (!this.endpoint) throw new Error('Endpoint URL must be entered in the Config sheet.');

    const scriptProps = PropertiesService.getScriptProperties();
    if (this.network === 'EVM') {
      this.apiKey = scriptProps.getProperty('ALCHEMY_API_KEY');
      if (!this.apiKey) throw new Error('ALCHEMY_API_KEY is missing in Script Properties.');
      this.apiEndpoint = [this.endpoint.replace(/\/$/, ''), this.apiKey].join('/');
      this.openSeaChainSlug = MultiChainNftTraitViewer.resolveOpenSeaChainSlug(this.endpoint);
    } else if (this.network === 'XRPL') {
      this.apiEndpoint = this.endpoint.replace(/\/$/, '');
    } else {
      throw new Error(`Unsupported network: ${this.network}. Please choose EVM or XRPL.`);
    }

    const lastRow = configSheet.getLastRow();
    this.displayTraits = lastRow >= 6 ? configSheet.getRange('A6:A' + lastRow).getValues().flat().filter(String) : [];
    if (this.displayTraits.length === 0) throw new Error('Please specify at least one trait to display starting from cell A6.');

    this.cacheSheet = this.ss.getSheetByName(MultiChainNftTraitViewer.CACHE_SHEET_NAME) ?? this.ss.insertSheet(MultiChainNftTraitViewer.CACHE_SHEET_NAME);
    if (this.cacheSheet.getRange(1, 1).getValue() === '') {
      this.cacheSheet.getRange('A1:D1').setValues([['Network', 'Owner Address', 'Contract/Issuer Address', 'Timestamp']]).setFontWeight('bold');
    }
  }

  readFromCache() {
    const targetAddress = this.contractAddress || 'ALL';
    const cacheData = this.cacheSheet.getDataRange().getValues();
    const cacheRow = cacheData.find(row => row[0] === this.network && row[1] === this.ownerAddress && row[2] === targetAddress);
    if (!cacheRow) return null;

    try {
      const chunks = cacheRow.slice(4).filter(String).map(JSON.parse);
      if (chunks.length === 0) return [];
      return chunks.flat();
    } catch (e) {
      this.ui.alert('Cache data is corrupted. Fetching from API.');
      return null;
    }
  }

  writeToCache(allOwnedNfts) {
    const targetAddress = this.contractAddress || 'ALL';
    const cacheData = this.cacheSheet.getDataRange().getValues();
    const cacheRowIndex = cacheData.findIndex(row => row[0] === this.network && row[1] === this.ownerAddress && row[2] === targetAddress);
    const targetRow = cacheRowIndex > -1 ? cacheRowIndex + 1 : this.cacheSheet.getLastRow() + 1;

    if (cacheRowIndex > -1) {
      const lastCol = this.cacheSheet.getLastColumn();
      if (lastCol > 0) {
        this.cacheSheet.getRange(targetRow, 1, 1, lastCol).clearContent();
      }
    }

    const chunkSize = 45000;
    const chunks = [];
    let candidate = [];
    for (const record of allOwnedNfts) {
      candidate.push(record);
      if (JSON.stringify(candidate).length >= chunkSize) {
        if (candidate.length === 1) {
          chunks.push(JSON.stringify(candidate));
          candidate = [];
        } else {
          chunks.push(JSON.stringify(candidate.slice(0, -1)));
          candidate = [record];
        }
      }
    }
    if (candidate.length > 0) chunks.push(JSON.stringify(candidate));

    const oversized = chunks.find(c => c.length > 49000);
    if (oversized) {
      throw new Error('NFT metadata is too large to cache for one item; consider fetching without cache.');
    }

    const rowData = [this.network, this.ownerAddress, targetAddress, new Date(), ...chunks];
    this.cacheSheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
  }

  fetchEvmFromApi() {
    let allOwnedNfts = [];
    let pageKey;
    const url = [
      `${this.apiEndpoint}/getNFTs`,
      `?owner=${this.ownerAddress}`,
      `&contractAddresses[]=${this.contractAddress}`,
      `&withMetadata=true`,
      (this.refresh ? `&refreshCache=true` : '')
    ].join('');

    do {
      const options = { 'method': 'get', 'contentType': 'application/json', 'muteHttpExceptions': true };
      const response = UrlFetchApp.fetch(`${url}${pageKey ? `&pageKey=${pageKey}` : ''}`, options);
      if (response.getResponseCode() !== 200) {
        throw new Error(`Alchemy API request failed [${response.getResponseCode()}]: ${response.getContentText()}`);
      }
      const data = JSON.parse(response.getContentText());
      if (data.ownedNfts) allOwnedNfts.push(...data.ownedNfts);
      pageKey = data.pageKey;
    } while (pageKey);

    this.writeToCache(allOwnedNfts);
    if (allOwnedNfts.length === 0) throw new Error('No NFTs found for the given criteria.');
    return allOwnedNfts;
  }

  // GASサーバー側で直接XRPLノードおよびIPFSメタデータを取得（中断・途中再開対応版）
  fetchXrplFromApi() {
    const startTime = Date.now();
    const TIME_LIMIT_MS = 4.5 * 60 * 1000; // 4.5分（270秒）で安全に中断

    // 1. キャッシュから既存の進行状況（中断データ）を読み込み
    let parsedNfts = this.refresh ? [] : this.readFromCache() || [];

    // キャッシュにデータがない（初回実行）場合、XRPLノードから一覧を取得
    if (parsedNfts.length === 0) {
      let rawNfts = [];
      let marker = null;
      let issuer = this.contractAddress || '';
      let taxon = null;

      if (issuer.includes('/')) {
        const parts = issuer.split('/');
        issuer = parts[0];
        taxon = parseInt(parts[1], 10);
      }

      do {
        const params = { account: this.ownerAddress, limit: 400 };
        if (marker) params.marker = marker;

        const payload = { method: 'account_nfts', params: [params] };
        const response = UrlFetchApp.fetch(this.apiEndpoint, {
          method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true
        });

        if (response.getResponseCode() !== 200) throw new Error(`XRPL Node Error: ${response.getContentText()}`);

        const data = JSON.parse(response.getContentText());
        if (data.result && data.result.account_nfts) {
          let filtered = data.result.account_nfts;
          if (issuer) filtered = filtered.filter(n => n.Issuer === issuer);
          if (taxon !== null && !isNaN(taxon)) filtered = filtered.filter(n => n.NFTokenTaxon === taxon);
          rawNfts.push(...filtered);
          marker = data.result.marker;
        } else {
          break;
        }
      } while (marker);

      // 初期状態の配列を作成 (status: 'pending')
      parsedNfts = rawNfts.map(nft => ({
        nftokenID: nft.NFTokenID || nft.nftokenID,
        URI: nft.URI,
        metadata: null,
        status: 'pending'
      }));
    }

    // 2. 未取得(pending)のメタデータのみ順次IPFSから取得
    let isInterrupted = false;

    for (let i = 0; i < parsedNfts.length; i++) {
      const nft = parsedNfts[i];

      // 取得完了済みのアイテムはスキップ
      if (nft.status === 'done') continue;

      if (nft.URI) {
        try {
          const uriStr = MultiChainNftTraitViewer.hexToUtf8(nft.URI);
          nft.metadata = MultiChainNftTraitViewer.fetchJsonFromIpfs(uriStr);
        } catch (e) {
          // エラー時はスキップして継続
        }
      }

      // 処理完了フラグを立てる
      nft.status = 'done';

      // 制限時間チェック（4.5分経過していたら中間保存してループ脱出）
      if (Date.now() - startTime > TIME_LIMIT_MS) {
        isInterrupted = true;
        break;
      }
    }

    // 3. 現在の進捗状態をキャッシュシートへ保存
    this.writeToCache(parsedNfts);

    // 中断された場合はユーザーへ通知
    if (isInterrupted) {
      const completedCount = parsedNfts.filter(n => n.status === 'done').length;
      throw new Error(`【タイムアウト回避のため一時停止】\n現在 ${completedCount} / ${parsedNfts.length} 件の取得が完了しました。\nもう一度「Fetch NFT Data (use cache)」を実行すると、続きから再開します。`);
    }

    if (parsedNfts.length === 0) throw new Error('No NFTs found for the given criteria.');
    return parsedNfts;
  }

  static extractTraits(attributes) {
    const traits = new Map();
    (attributes ?? []).forEach(attr => {
      if (attr.trait_type && attr.value !== undefined) {
        traits.set(String(attr.trait_type).toLowerCase(), String(attr.value));
      }
    });
    return traits;
  }

  static normalizeImageUrl(imageUrl) {
    if (imageUrl && imageUrl.startsWith('ipfs://')) {
      return imageUrl.replace('ipfs://', 'https://gateway.pinata.cloud/ipfs/');
    }
    return imageUrl || '';
  }

  // 利用するIPFSゲートウェイのリスト（優先度順）
  static get IPFS_GATEWAYS() {
    return [
      'https://gateway.pinata.cloud/ipfs/',
      'https://nftstorage.link/ipfs/',
      'https://dweb.link/ipfs/',
      'https://cloudflare-ipfs.com/ipfs/',
      'https://ipfs.filebase.io/ipfs/'
    ];
  }

  // ipfs:// URLを各種ゲートウェイURLに変換
  static getGatewayUrls(uri) {
    if (!uri) return [];
    if (uri.startsWith('ipfs://')) {
      const cid = uri.replace('ipfs://', '').replace(/^ipfs\//, '');
      return MultiChainNftTraitViewer.IPFS_GATEWAYS.map(gw => `${gw}${cid}`);
    }
    return [uri];
  }

  // GASからブロックされずにJSONメタデータを取得するためのマルチゲートウェイFetch
  static fetchJsonFromIpfs(uri) {
    const urls = MultiChainNftTraitViewer.getGatewayUrls(uri);
    for (const url of urls) {
      try {
        const res = UrlFetchApp.fetch(url, {
          muteHttpExceptions: true,
          headers: { 'User-Agent': 'Mozilla/5.0' }
        });
        if (res.getResponseCode() === 200) {
          try { return JSON.parse(res.getContentText()); } catch (e) { return null; }
        }
      } catch (e) {
        // 次のゲートウェイへフォールバック
      }
    }
    return null;
  }

  getNftIdAndImage(nft) {
    if (this.network === 'EVM') {
      const id = nft.id?.tokenId;
      const imageUrl = (nft.media && nft.media[0]) ? nft.media[0].gateway : nft.tokenUri?.gateway;
      return { id, imageUrl: MultiChainNftTraitViewer.normalizeImageUrl(imageUrl) };
    }
    const id = nft.nftokenID;
    const imageUrl = nft.metadata?.image || nft.metadata?.image_url || '';
    return { id, imageUrl: MultiChainNftTraitViewer.normalizeImageUrl(imageUrl) };
  }

  groupBy(allOwnedNfts, displayTraits) {
    const groupedNfts = new Map();

    allOwnedNfts.forEach(nft => {
      const nftTraits = MultiChainNftTraitViewer.extractTraits(nft.metadata?.attributes);
      const { id, imageUrl } = this.getNftIdAndImage(nft);

      const groupValues = displayTraits.map(trait => nftTraits.get(trait.toLowerCase()) ?? '');
      const groupKey = groupValues.join('-');

      if (!groupedNfts.has(groupKey)) {
        groupedNfts.set(groupKey, { values: groupValues, nfts: [] });
      }
      groupedNfts.get(groupKey).nfts.push({ id, imageUrl });
    });

    return groupedNfts;
  }

  getExplorerUrl(id) {
    if (this.network === 'EVM') {
      return `https://opensea.io/assets/${this.openSeaChainSlug}/${this.contractAddress}/${BigInt(id).toString(10)}`;
    }
    return `https://bithomp.com/nft/${id}`;
  }

  getRecords(groupedNfts, maxImages, imageHeaderStart) {
    const rowIndexOffset = 2;
    return [...groupedNfts.values()].map((group, rowIndex) => {
      group.nfts.sort((a, b) => {
        if (this.network === 'EVM') {
          try {
            const idA = BigInt(a.id); const idB = BigInt(b.id);
            return (idA < idB) ? -1 : (idA > idB) ? 1 : 0;
          } catch (e) { return String(a.id).localeCompare(String(b.id)); }
        }
        return String(a.id).localeCompare(String(b.id));
      });

      let countFormula = 0;
      if (maxImages > 0) {
        const imageStartColLetter = MultiChainNftTraitViewer.columnToLetter(imageHeaderStart + 1);
        const imageEndColLetter = MultiChainNftTraitViewer.columnToLetter(imageHeaderStart + maxImages);
        const countaRange = `${imageStartColLetter}${rowIndex + rowIndexOffset}:${imageEndColLetter}${rowIndex + rowIndexOffset}`;
        countFormula = `=COUNTA(${countaRange})`;
      }

      return [
        countFormula,
        this.ownerAddress,
        this.contractAddress || 'ALL',
        ...group.values,
        ...Array(maxImages).fill('').map((_, colIndex) => (colIndex < group.nfts.length) ?
          `=HYPERLINK("${this.getExplorerUrl(group.nfts[colIndex].id)}", IMAGE("${group.nfts[colIndex].imageUrl}", 1))` :
          ''
        )
      ];
    });
  }

  getUsableCache() {
    if (!this.useCache || this.refresh) return null;
    return this.readFromCache();
  }

  build() {
    try {
      const cached = this.getUsableCache();
      if (cached !== null) {
        this.ui.alert('Using cached data.');
        if (cached.length === 0) throw new Error('No NFTs found in cache.');
        this.renderSheet(cached);
        return;
      }

      this.ui.alert(`Fetching NFTs via ${this.network}... This may take a moment.`);
      const allOwnedNfts = (this.network === 'XRPL') ? this.fetchXrplFromApi() : this.fetchEvmFromApi();
      this.renderSheet(allOwnedNfts);

    } catch (error) { this.ui.alert('Error: ' + error.message); }
  }

  renderSheet(allOwnedNfts) {
    const groupedNfts = this.groupBy(allOwnedNfts, this.displayTraits);
    const maxImages = groupedNfts.size > 0 ? Math.max(...[...groupedNfts.values()].map(group => group.nfts.length)) : 0;

    const headersFixed = ['Count', 'Owner Address', 'Contract/Issuer Address'];
    const headersImages = Array(maxImages).fill('').map((_, k) => `Image ${k + 1}`);
    const headers = [...headersFixed, ...this.displayTraits, ...headersImages];
    const rows = this.getRecords(groupedNfts, maxImages, this.displayTraits.length + headersFixed.length);

    const ownerShort = this.ownerAddress.slice(-6);
    const contractShort = (this.contractAddress || 'ALL').replace('/', '_').slice(-10);
    const dataSheetName = `${this.network}_${ownerShort}_${contractShort}`;

    const dataSheet = this.ss.getSheetByName(dataSheetName) ?? this.ss.insertSheet(dataSheetName);
    dataSheet.clear();
    dataSheet.getFilter()?.remove();

    dataSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    if (rows.length > 0) dataSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

    dataSheet.setRowHeights(2, Math.max(dataSheet.getLastRow() - 1, 1), 32);
    dataSheet.autoResizeColumns(headersFixed.length + 1, this.displayTraits.length);
    if (maxImages > 0) {
      dataSheet.setColumnWidths(headersFixed.length + this.displayTraits.length + 1, maxImages, 32);
    }

    dataSheet.hideColumns(2, 2);
    dataSheet.getDataRange().createFilter();

    dataSheet.activate();
    this.ui.alert(`Success! ${groupedNfts.size} groups written to '${dataSheetName}' sheet.`);
  }
}

// ==========================================
// メニュー呼び出し用トップレベルラッパー関数
// ==========================================
function setupConfigSheet() { MultiChainNftTraitViewer.setupConfigSheet(); }
function fetchNftDataWithCache() { MultiChainNftTraitViewer.fetchNftDataWithCache(); }
function fetchNftDataWithoutCache() { MultiChainNftTraitViewer.fetchNftDataWithoutCache(); }
function fetchNftDataWithRefreshedCache() { MultiChainNftTraitViewer.fetchNftDataWithRefreshedCache(); }

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('NFT Trait Viewer')
    .addItem('1. Setup Config Sheet', 'setupConfigSheet')
    .addSeparator()
    .addItem('2. Fetch NFT Data (use cache)', 'fetchNftDataWithCache')
    .addItem('3. Fetch NFT Data (no cache)', 'fetchNftDataWithoutCache')
    .addItem('4. Fetch NFT Data (refreshed cache)', 'fetchNftDataWithRefreshedCache')
    .addToUi();
}
/**
 * @OnlyCurrentDoc
 */

class MultiChainNftTraitViewer {
  static get CONFIG_SHEET_NAME() { return 'Config'; }
  static get CACHE_SHEET_NAME() { return 'Cache'; }

  static setupConfigSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(MultiChainNftTraitViewer.CONFIG_SHEET_NAME) ?? ss.insertSheet(MultiChainNftTraitViewer.CONFIG_SHEET_NAME);
    sheet.clear();

    sheet.getRange('A1:B4').setValues([
      ['Network (EVM / XRPL):', 'EVM'],
      ['Endpoint (Alchemy or Bithomp):', 'https://eth-mainnet.g.alchemy.com/v2/'],
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
      '   - EVM: https://eth-mainnet.g.alchemy.com/v2/',
      '   - XRPL: https://bithomp.com/api/v2/nfts',
      '3. Enter Owner Address (B3), Contract/Issuer Address (B4, XRPL can be issuer/taxon), and Traits starting from cell A6.'
    ].join('\n'));
  }

  static fetchNftDataWithRefreshedCache() {
    (new MultiChainNftTraitViewer({
      ss: SpreadsheetApp.getActiveSpreadsheet(),
      ui: SpreadsheetApp.getUi(),
      useCache: false,
      refresh: true
    })).build();
  }

  static fetchNftDataWithCache() {
    (new MultiChainNftTraitViewer({
      ss: SpreadsheetApp.getActiveSpreadsheet(),
      ui: SpreadsheetApp.getUi(),
      useCache: true
    })).build();
  }

  static fetchNftDataWithoutCache() {
    (new MultiChainNftTraitViewer({
      ss: SpreadsheetApp.getActiveSpreadsheet(),
      ui: SpreadsheetApp.getUi(),
      useCache: false
    })).build();
  }

  static columnToLetter(column) {
    const cp = (e) => ((e) => e > 57 ? e - 39 : e)(e.codePointAt(0));
    const cs = (e) => String.fromCodePoint(e + 'A'.codePointAt(0));
    const th = (e) => Array(e).fill('').reduce((a, _, k) => a + 26 ** (k + 1), 0);
    for (let k = 0; k < 3; k++) {
      if (column <= th(k + 1)) {
        return [...(column - th(k) - 1).toString(26).padStart(k + 1, '0')].map(e => cs(cp(e) - cp('0'))).join('');
      }
    }
    throw new Error('Column number exceeds the supported range.');
  }

  constructor({ ss, ui, useCache, refresh = false }) {
    Object.assign(this, { ss, ui, useCache, refresh });
    try { this.init(); } catch (error) { this.ui.alert('Error: ' + error.message); }
  }

  init() {
    const configSheet = this.ss.getSheetByName(MultiChainNftTraitViewer.CONFIG_SHEET_NAME);
    if (!configSheet) throw new Error(`Sheet "${MultiChainNftTraitViewer.CONFIG_SHEET_NAME}" not found. Please run "Setup Config Sheet" first.`);

    [this.network, this.endpoint, this.ownerAddress, this.contractAddress] = configSheet.getRange('B1:B4').getValues().flat();
    this.network = (this.network || 'EVM').toUpperCase();

    if (!this.ownerAddress) throw new Error('Owner Address must be entered in the Config sheet.');
    if (this.network === 'EVM' && !this.contractAddress) throw new Error('Contract Address must be entered in the Config sheet for EVM.');
    if (!this.endpoint) throw new Error('Endpoint URL must be entered in the Config sheet.');

    const scriptProps = PropertiesService.getScriptProperties();
    if (this.network === 'EVM') {
      this.apiKey = scriptProps.getProperty('ALCHEMY_API_KEY');
      if (!this.apiKey) throw new Error('ALCHEMY_API_KEY is missing in Script Properties.');
      this.apiEndpoint = [this.endpoint.replace(/\/$/, ''), this.apiKey].join('/');
    } else if (this.network === 'XRPL') {
      this.apiKey = scriptProps.getProperty('BITHOMP_API_KEY');
      if (!this.apiKey) throw new Error('BITHOMP_API_KEY is missing in Script Properties.');
      this.apiEndpoint = this.endpoint.replace(/\/$/, '');
    } else {
      throw new Error(`Unsupported network: ${this.network}. Please choose EVM or XRPL.`);
    }

    this.displayTraits = configSheet.getRange('A6:A' + configSheet.getLastRow()).getValues().flat().filter(String);
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
      if (lastCol > 4) {
        this.cacheSheet.getRange(targetRow, 5, 1, lastCol - 4).clearContent();
      }
    }

    const chunkSize = 45000;
    const chunks = [];
    let candidate = [];
    for (const record of allOwnedNfts) {
      candidate.push(record);
      if (JSON.stringify(candidate).length >= chunkSize) {
        chunks.push(JSON.stringify(candidate.slice(0, -1)));
        candidate = [record];
      }
    }
    if (candidate.length > 0) chunks.push(JSON.stringify(candidate));
    const rowData = [this.network, this.ownerAddress, targetAddress, new Date(), ...chunks];
    this.cacheSheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
  }

  fetchFromApi() {
    this.ui.alert(`Fetching NFTs via ${this.network}... This may take a moment.`);
    return this.network === 'EVM' ? this.fetchEvmFromApi() : this.fetchXrplFromApi();
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

  fetchXrplFromApi() {
    let allOwnedNfts = [];
    let marker = null;

    // Issuer と Taxon の分離ロジック
    let issuer = this.contractAddress || '';
    let taxon = null;
    if (issuer.includes('/')) {
      const parts = issuer.split('/');
      issuer = parts[0];
      taxon = parts[1];
    }

    do {
      let url = `${this.apiEndpoint}?owner=${this.ownerAddress}`;
      if (issuer) url += `&issuer=${issuer}`;
      if (taxon) url += `&taxon=${taxon}`;
      if (marker) url += `&marker=${marker}`;

      const options = {
        'method': 'get',
        'headers': { 'x-bithomp-token': this.apiKey },
        'muteHttpExceptions': true
      };

      const response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() !== 200) {
        throw new Error(`Bithomp API request failed [${response.getResponseCode()}]: ${response.getContentText()}`);
      }

      const data = JSON.parse(response.getContentText());
      if (data.nfts) allOwnedNfts.push(...data.nfts);
      marker = data.marker;
    } while (marker);

    this.writeToCache(allOwnedNfts);
    if (allOwnedNfts.length === 0) throw new Error('No NFTs found for the given criteria.');
    return allOwnedNfts;
  }

  fetch() {
    if (this.useCache) {
      const cachedNfts = this.readFromCache();
      if (cachedNfts !== null) {
        this.ui.alert('Using cached data.');
        if (cachedNfts.length === 0) throw new Error('No NFTs found in cache.');
        return cachedNfts;
      }
    }
    return this.fetchFromApi();
  }

  groupBy(allOwnedNfts, displayTraits) {
    const groupedNfts = new Map();

    allOwnedNfts.forEach(nft => {
      const nftTraits = new Map();
      let id = '';
      let imageUrl = '';

      if (this.network === 'EVM') {
        const attributes = nft.metadata?.attributes ?? [];
        attributes.forEach(attr => {
          if (attr.trait_type && attr.value !== undefined) {
            nftTraits.set(String(attr.trait_type).toLowerCase(), String(attr.value));
          }
        });
        id = nft.id?.tokenId;
        imageUrl = (nft.media && nft.media[0]) ? nft.media[0].gateway : nft.tokenUri?.gateway;
      } else {
        const attributes = nft.metadata?.attributes ?? [];
        attributes.forEach(attr => {
          if (attr.trait_type && attr.value !== undefined) {
            nftTraits.set(String(attr.trait_type).toLowerCase(), String(attr.value));
          }
        });
        id = nft.nftokenID;
        // Public IPFS または HTTP 画像URLを取得
        imageUrl = nft.metadata?.image || nft.metadata?.image_url || '';
      }

      if (imageUrl && imageUrl.startsWith('ipfs://')) {
        imageUrl = imageUrl.replace('ipfs://', 'https://ipfs.io/ipfs/');
      }

      const groupValues = displayTraits.map(trait => nftTraits.get(trait.toLowerCase()) ?? '');
      const groupKey = groupValues.join('-');

      if (!groupedNfts.has(groupKey)) {
        groupedNfts.set(groupKey, { values: groupValues, nfts: [] });
      }
      groupedNfts.get(groupKey).nfts.push({ id, imageUrl });
    });

    return groupedNfts;
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
        } else {
          return String(a.id).localeCompare(String(b.id));
        }
      });

      const imageStartColLetter = MultiChainNftTraitViewer.columnToLetter(imageHeaderStart + 1);
      const imageEndColLetter = MultiChainNftTraitViewer.columnToLetter(imageHeaderStart + maxImages);
      const countaRange = `${imageStartColLetter}${rowIndex + rowIndexOffset}:${imageEndColLetter}${rowIndex + rowIndexOffset}`;

      const getExplorerUrl = (id) => {
        if (this.network === 'EVM') {
          return `https://opensea.io/assets/ethereum/${this.contractAddress}/${BigInt(id).toString(10)}`;
        } else {
          return `https://bithomp.com/nft/${id}`;
        }
      };

      return [
        `=COUNTA(${countaRange})`,
        this.ownerAddress,
        this.contractAddress || 'ALL',
        ...group.values,
        ...Array(maxImages).fill('').map((_, colIndex) => (colIndex < group.nfts.length) ?
          `=HYPERLINK("${getExplorerUrl(group.nfts[colIndex].id)}", IMAGE("${group.nfts[colIndex].imageUrl}", 1))` :
          ''
        )
      ];
    });
  }

  build() {
    try {
      const allOwnedNfts = this.fetch();
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
    } catch (error) { this.ui.alert('Error: ' + error.message); }
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('NFT Trait Viewer')
    .addItem('1. Setup Config Sheet', 'MultiChainNftTraitViewer.setupConfigSheet')
    .addSeparator()
    .addItem('2. Fetch NFT Data (use cache)', 'MultiChainNftTraitViewer.fetchNftDataWithCache')
    .addItem('3. Fetch NFT Data (no cache)', 'MultiChainNftTraitViewer.fetchNftDataWithoutCache')
    .addItem('4. Fetch NFT Data (refreshed cache)', 'MultiChainNftTraitViewer.fetchNftDataWithRefreshedCache')
    .addToUi();
}

/**
 * @OnlyCurrentDoc
 */

class NftTraitViewer {
  static get CONFIG_SHEET_NAME() { return 'Config'; }

  static setupConfigSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(NftTraitViewer.CONFIG_SHEET_NAME) ?? ss.insertSheet(NftTraitViewer.CONFIG_SHEET_NAME);
    sheet.clear();
    sheet.getRange('A1').setValue('Owner Address:');
    sheet.getRange('A2').setValue('Contract Address:');
    sheet.getRange('A4').setValue('Traits to Display (one per cell below):');
    sheet.getRange('A1:A2').setFontWeight('bold');
    sheet.getRange('A4').setFontWeight('bold');
    SpreadsheetApp.getUi().alert('Config sheet has been set up. Please enter the addresses in cells B1 and B2, and the traits you want to display starting from cell A5.');
  }
  static fetchNftData() {
    (new NftTraitViewer({
      ss: SpreadsheetApp.getActiveSpreadsheet(),
      ui: SpreadsheetApp.getUi(),
      apiKey: PropertiesService.getScriptProperties().getProperty('ALCHEMY_API_KEY')
    })).build();
  }
  static columnToLetter(column) {
    let temp, letter = '';
    while (column > 0) {
      temp = (column - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      column = (column - temp - 1) / 26;
    }
    return letter;
  }
  constructor({ ss, ui, apiKey }) {
    Object.assign(this, { ss, ui, apiKey });
    try { this.init(); } catch (error) { this.ui.alert('Error: ' + error.message); }
  }

  init() {
    if (!this.apiKey || !this.apiKey.startsWith('https://')) throw new Error('Alchemy API Key is invalid or missing. Please set it in the script properties.');
    const configSheet = this.ss.getSheetByName(NftTraitViewer.CONFIG_SHEET_NAME);
    if (!configSheet) throw new Error(`Sheet "${NftTraitViewer.CONFIG_SHEET_NAME}" not found. Please run "Setup Config Sheet" first.`);
    this.ownerAddress = configSheet.getRange('B1').getValue();
    this.contractAddress = configSheet.getRange('B2').getValue();
    if (!this.ownerAddress || !this.contractAddress) throw new Error('Owner Address and Contract Address must be entered in the Config sheet.');
    this.displayTraits = configSheet.getRange('A5:A' + configSheet.getLastRow()).getValues().flat().filter(String);
    if (this.displayTraits.length === 0) throw new Error('Please specify at least one trait to display in the Config sheet (starting from cell A5).');
  }
  fetch() {
    let allOwnedNfts = [];
    let pageKey;
    const url = `${this.apiKey}/getNFTs?owner=${this.ownerAddress}&contractAddresses[]=${this.contractAddress}&withMetadata=true`;
    do {
      const options = { 'method': 'get', 'contentType': 'application/json', 'muteHttpExceptions': true };
      const response = UrlFetchApp.fetch(`${url}${pageKey ? `&pageKey=${pageKey}` : ''}`, options);
      const responseCode = response.getResponseCode();
      const responseBody = response.getContentText();
      if (responseCode !== 200) throw new Error(`API request failed with status ${responseCode}. Response: ${responseBody}`);
      const data = JSON.parse(responseBody);
      if (data.ownedNfts) allOwnedNfts.push(...data.ownedNfts);
      pageKey = data.pageKey;
    } while (pageKey);
    if (allOwnedNfts.length === 0) throw new Error('No NFTs found for the given address and contract.');
    return allOwnedNfts;
  }
  groupBy(allOwnedNfts, displayTraits) {
    const groupedNfts = new Map();
    allOwnedNfts.forEach(nft => {
      const nftTraits = (nft.metadata?.attributes ?? []).filter(attr => attr.trait_type && attr.value).reduce((map, attr) => map.set(attr.trait_type, String(attr.value)), new Map());
      const groupValues = displayTraits.map(trait => nftTraits.get(trait) ?? '');
      const groupKey = groupValues.join('-');
      if (!groupedNfts.has(groupKey)) { groupedNfts.set(groupKey, { values: groupValues, nfts: [] }); }
      const imageUrl = (nft.media && nft.media[0]) ? nft.media[0].gateway : nft.tokenUri?.gateway;
      groupedNfts.get(groupKey).nfts.push({ id: nft.id?.tokenId, imageUrl: imageUrl });
    });
    return groupedNfts;
  }
  getRecords(groupedNfts, maxImages, imageHeaderStart) {
    const rowIndexOffset = 2;
    return [...groupedNfts.values()].map((group, rowIndex) => {
      const imageStartColLetter = NftTraitViewer.columnToLetter(imageHeaderStart + 1);
      const imageEndColLetter = NftTraitViewer.columnToLetter(imageHeaderStart + maxImages);
      const countaRange = `${imageStartColLetter}${rowIndex + rowIndexOffset}:${imageEndColLetter}${rowIndex + rowIndexOffset}`;
      const row = [`=COUNTA(${countaRange})`, this.ownerAddress, this.contractAddress, ...group.values];
      group.nfts.sort((a, b) => {
        try {
          const idA = BigInt(a.id); const idB = BigInt(b.id);
          return (idA < idB) ? -1 : (idA > idB) ? 1 : 0;
        } catch (e) { return String(a.id).localeCompare(String(b.id)); }
      });
      for (let i = 0; i < maxImages; i++) {
        if (i < group.nfts.length) {
          const tokenId = BigInt(group.nfts[i].id).toString(10);
          const imageUrl = group.nfts[i].imageUrl;
          const openSeaUrl = `https://opensea.io/assets/ethereum/${this.contractAddress}/${tokenId}`;
          row.push(`=HYPERLINK("${openSeaUrl}", IMAGE("${imageUrl}", 1))`);
        } else { row.push(''); }
      }
      return row;
    });
  }
  build() {
    try {
      this.ui.alert('Fetching all NFTs... This may take a moment and involve multiple API calls.');
      const allOwnedNfts = this.fetch();
      const groupedNfts = this.groupBy(allOwnedNfts, this.displayTraits);
      const maxImages = Math.max(...[...groupedNfts.values()].map(group => group.nfts.length));

      const headers = ['Count', 'Owner Address', 'Contract Address', ...this.displayTraits];
      const rows = this.getRecords(groupedNfts, maxImages, headers.length);
      for (let i = 1; i <= maxImages; i++) headers.push(`Image ${i}`);

      const dataSheetName = [this.ownerAddress.slice(-6), this.contractAddress.slice(-6)].join('/');
      const dataSheet = this.ss.getSheetByName(dataSheetName) ?? this.ss.insertSheet(dataSheetName);
      dataSheet.clear();

      dataSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      if (rows.length > 0) dataSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

      dataSheet.hideColumns(2, 2);
      dataSheet.getDataRange().createFilter();

      if (rows.length > 0) {
        for (let i = 2; i <= rows.length + 1; i++) { dataSheet.setRowHeight(i, 32); }
        if (this.displayTraits.length > 0) dataSheet.autoResizeColumns(headers.length - maxImages + 1 - this.displayTraits.length, this.displayTraits.length);
        if (maxImages > 0) { for (let i = 0; i < maxImages; i++) { dataSheet.setColumnWidth(headers.length - maxImages + 1 + i, 32); } }
      }

      dataSheet.activate();
      this.ui.alert(`Success! ${groupedNfts.size} groups have been written to the '${dataSheetName}' sheet.`);
    } catch (error) { this.ui.alert('Error: ' + error.message); }
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('NFT Viewer')
    .addItem('1. Setup Config Sheet', 'NftTraitViewer.setupConfigSheet')
    .addSeparator()
    .addItem('2. Fetch NFT Data', 'NftTraitViewer.fetchNftData')
    .addToUi();
}

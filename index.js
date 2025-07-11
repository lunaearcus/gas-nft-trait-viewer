/**
 * @OnlyCurrentDoc
 */

class NftTraitViewer {
  static get CONFIG_SHEET_NAME() { return 'Config'; }

  static setupConfigSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(NftTraitViewer.CONFIG_SHEET_NAME) ?? ss.insertSheet(NftTraitViewer.CONFIG_SHEET_NAME);
    sheet.clear();
    sheet.getRange('A1:B3').setValues([
      ['AlchemyEndpoint:', 'https://eth-mainnet.g.alchemy.com/v2/'],
      ['Owner Address:', ''],
      ['Contract Address:', '']
    ]);
    sheet.getRange('A4').setValue('Traits to Display (one per cell below):');
    sheet.getRange('A1:A4').setFontWeight('bold');
    SpreadsheetApp.getUi().alert([
      'Config sheet has been set up.',
      'Please enter the addresses in cells B2 and B3, and the traits you want to display starting from cell A5.'
    ].join('\n'));
  }
  static fetchNftData() {
    (new NftTraitViewer({
      ss: SpreadsheetApp.getActiveSpreadsheet(),
      ui: SpreadsheetApp.getUi(),
      apiKey: PropertiesService.getScriptProperties().getProperty('ALCHEMY_API_KEY')
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
    throw new Error('Column number exceeds the supported range (up to 18272 for 3-letter columns).');
  }
  constructor({ ss, ui, apiKey }) {
    Object.assign(this, { ss, ui, apiKey });
    try { this.init(); } catch (error) { this.ui.alert('Error: ' + error.message); }
  }

  init() {
    if (!this.apiKey) throw new Error('Alchemy API Key is missing. Please set it in the script properties.');
    const configSheet = this.ss.getSheetByName(NftTraitViewer.CONFIG_SHEET_NAME);
    if (!configSheet) throw new Error(`Sheet "${NftTraitViewer.CONFIG_SHEET_NAME}" not found. Please run "Setup Config Sheet" first.`);
    [this.alchemyEndpoint, this.ownerAddress, this.contractAddress] = configSheet.getRange('B1:B3').getValues().flat();
    if (!this.alchemyEndpoint || !this.ownerAddress || !this.contractAddress) throw new Error('AlchemyEndpoint, Owner Address and Contract Address must be entered in the Config sheet.');
    this.displayTraits = configSheet.getRange('A5:A' + configSheet.getLastRow()).getValues().flat().filter(String);
    if (this.displayTraits.length === 0) throw new Error('Please specify at least one trait to display in the Config sheet (starting from cell A5).');
    this.apiEndpoint = [this.alchemyEndpoint.replace(/\/$/, ''), this.apiKey].join('/');
  }
  fetch() {
    let allOwnedNfts = [];
    let pageKey;
    const url = `${this.apiEndpoint}/getNFTs?owner=${this.ownerAddress}&contractAddresses[]=${this.contractAddress}&withMetadata=true`;
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
      group.nfts.sort((a, b) => {
        try {
          const idA = BigInt(a.id); const idB = BigInt(b.id);
          return (idA < idB) ? -1 : (idA > idB) ? 1 : 0;
        } catch (e) { return String(a.id).localeCompare(String(b.id)); }
      });
      const imageStartColLetter = NftTraitViewer.columnToLetter(imageHeaderStart + 1);
      const imageEndColLetter = NftTraitViewer.columnToLetter(imageHeaderStart + maxImages);
      const countaRange = `${imageStartColLetter}${rowIndex + rowIndexOffset}:${imageEndColLetter}${rowIndex + rowIndexOffset}`;
      const getOpenSeaUrl = (id) => `https://opensea.io/assets/ethereum/${this.contractAddress}/${BigInt(id).toString(10)}`;
      return [
        `=COUNTA(${countaRange})`, this.ownerAddress, this.contractAddress,
        ...group.values,
        ...Array(maxImages).fill('').map((_, colIndex) => (colIndex < group.nfts.length) ?
          `=HYPERLINK("${getOpenSeaUrl(group.nfts[colIndex].id)}", IMAGE("${group.nfts[colIndex].imageUrl}", 1))` :
          ''
        )
      ];
    });
  }
  build() {
    try {
      this.ui.alert('Fetching all NFTs... This may take a moment and involve multiple API calls.');
      const allOwnedNfts = this.fetch();
      const groupedNfts = this.groupBy(allOwnedNfts, this.displayTraits);
      const maxImages = Math.max(...[...groupedNfts.values()].map(group => group.nfts.length));

      // Prepare the headers and rows for the data sheet
      const headersFixed = ['Count', 'Owner Address', 'Contract Address'];
      const headersImages = Array(maxImages).fill('').map((_, k) => `Image ${k + 1}`);
      const headers = [...headersFixed, ...this.displayTraits, ...headersImages];
      const rows = this.getRecords(groupedNfts, maxImages, this.displayTraits.length + headersFixed.length);

      // Create or clear the data sheet
      const dataSheetName = [this.ownerAddress.slice(-6), this.contractAddress.slice(-6)].join('/');
      const dataSheet = this.ss.getSheetByName(dataSheetName) ?? this.ss.insertSheet(dataSheetName);
      dataSheet.clear();
      dataSheet.getFilter()?.remove();

      // Set the headers and rows in the data sheet
      dataSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      if (rows.length > 0) dataSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

      // Set column widths and row heights
      dataSheet.setRowHeights(2, dataSheet.getLastRow(), 32);
      dataSheet.autoResizeColumns(headersFixed.length + 1, this.displayTraits.length);
      dataSheet.setColumnWidths(headersFixed.length + this.displayTraits.length + 1, maxImages, 32);

      // Set the first three columns to be hidden and create a filter
      dataSheet.hideColumns(2, 2);
      dataSheet.getDataRange().createFilter();

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

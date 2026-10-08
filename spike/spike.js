/* Essai du complément « Action ! » sur Excel 2021 — vérifie en vrai les points incertains (plan 3, tâche 1). */
const SUPABASE_URL = 'https://ibyigjqviaoaqnjjnxsp.supabase.co'
const SUPABASE_KEY = 'sb_publishable_jFkXsqvZpSM1bRNc25pjOw_GlrJV4Pk'
const SHEET = 'Essai'
const TABLE = 'Essai'
const LISTS = 'ListesEssai'
const HEADERS = ['ID', 'N°', 'Photo', 'Action', 'Pièce', 'Échéance', 'Statut']
const ROOMS = ['Cuisine', 'Entrée', 'Séjour']
const logArea = document.getElementById('log')
let eventCount = 0
let supabaseClient = null

function log(id, text) {
  logArea.value += `[${id}] ${text}\n`
  logArea.scrollTop = logArea.scrollHeight
}
const describeError = (e) => `${e && e.code ? e.code + ' ' : ''}${e && e.message ? e.message : String(e)}`
async function step(id, fn) {
  try {
    await fn()
  } catch (e) {
    log(id, `ERREUR ${describeError(e)}`)
  }
}
async function check(id, label, fn) {
  try {
    const r = await fn()
    log(id, `${label} : OK${r !== undefined ? ' — ' + r : ''}`)
  } catch (e) {
    log(id, `${label} : ERREUR ${describeError(e)}`)
  }
}
const on = (id, fn) => document.getElementById(id).addEventListener('click', () => step(id, fn))
const cm = (n) => Math.round(((n * 72) / 2.54) * 100) / 100

function testJpeg(w, h, label) {
  const canvas = document.createElement('canvas')
  canvas.width = w
  canvas.height = h
  const g = canvas.getContext('2d')
  g.fillStyle = '#8A6D3B'
  g.fillRect(0, 0, w, h)
  g.fillStyle = '#FFFFFF'
  g.font = `${Math.round(h / 4)}px sans-serif`
  g.fillText(label, w / 10, h / 2)
  const url = canvas.toDataURL('image/jpeg', 0.8)
  return url.slice(url.indexOf(',') + 1)
}

function idbOpen() {
  return new Promise((resolve, reject) => {
    const r = indexedDB.open('action-spike', 1)
    r.onupgradeneeded = () => r.result.createObjectStore('kv')
    r.onsuccess = () => resolve(r.result)
    r.onerror = () => reject(r.error)
  })
}
async function idbGet() {
  const db = await idbOpen()
  return new Promise((resolve, reject) => {
    const r = db.transaction('kv').objectStore('kv').get('stamp')
    r.onsuccess = () => resolve(r.result)
    r.onerror = () => reject(r.error)
  })
}
async function idbPut(value) {
  const db = await idbOpen()
  return new Promise((resolve, reject) => {
    const tx = db.transaction('kv', 'readwrite')
    tx.objectStore('kv').put(value, 'stamp')
    tx.oncomplete = () => resolve()
    tx.onerror = () => reject(tx.error)
  })
}
function client() {
  if (!supabaseClient) {
    supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY, {
      auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: false },
    })
  }
  return supabaseClient
}

Office.onReady(async (info) => {
  log('0', `Office prêt : hôte ${info.host}, plateforme ${info.platform}`)
  try {
    log('7', `localStorage au démarrage : ${localStorage.getItem('action-spike') || '(rien)'}`)
  } catch (e) {
    log('7', `localStorage ERREUR ${describeError(e)}`)
  }
  await step('7', async () => log('7', `IndexedDB au démarrage : ${(await idbGet()) || '(rien)'}`))
  if (window.supabase && SUPABASE_URL.startsWith('https://')) {
    await step('8', async () => {
      const { data, error } = await client().auth.getSession()
      if (error) throw error
      log('8', `session au démarrage : ${data.session ? 'présente (' + data.session.user.email + ')' : 'aucune'}`)
    })
  } else {
    log('8', 'bibliothèque Supabase non chargée ou adresse absente')
  }
  await step('9', () =>
    Excel.run(async (ctx) => {
      const s = ctx.workbook.settings.getItemOrNullObject('action.essai')
      s.load('value')
      await ctx.sync()
      log('9', `réglage du classeur au démarrage : ${s.isNullObject ? '(aucun)' : s.value}`)
    }),
  )
})

on('b1', async () => {
  const d = Office.context.diagnostics
  log('1', `Excel ${d.version} (${d.platform})`)
  for (const v of ['1.9', '1.10', '1.11', '1.12', '1.13', '1.14', '1.15', '1.16']) {
    log('1', `ExcelApi ${v} : ${Office.context.requirements.isSetSupported('ExcelApi', v) ? 'oui' : 'non'}`)
  }
  for (const [set, v] of [['DialogApi', '1.1'], ['SharedRuntime', '1.1'], ['AddinCommands', '1.1']]) {
    log('1', `${set} ${v} : ${Office.context.requirements.isSetSupported(set, v) ? 'oui' : 'non'}`)
  }
  log('1', `Navigateur : ${navigator.userAgent}`)
})

on('b2', () =>
  Excel.run(async (ctx) => {
    for (const name of [SHEET, LISTS]) {
      const old = ctx.workbook.worksheets.getItemOrNullObject(name)
      await ctx.sync()
      if (!old.isNullObject) {
        old.delete()
        await ctx.sync()
      }
    }
    const ws = ctx.workbook.worksheets.add(SHEET)
    ws.activate()
    const rows = Array.from({ length: 30 }, (_, i) => [
      `id-${i + 1}`, i + 1, '', `Action ${i + 1}\nCommentaire ${i + 1}`, ROOMS[i % 3], 46300 + i, i % 2 ? 'Fait' : 'À faire',
    ])
    ws.getRange('A1:G1').values = [HEADERS]
    ws.getRange('A2:G31').values = rows
    const t = ws.tables.add('Essai!A1:G31', true)
    t.name = TABLE
    await ctx.sync()
    log('2', 'tableau créé')
    await check('2', 'S3 colonne ID masquée', async () => {
      ws.getRange('A:A').columnHidden = true
      await ctx.sync()
    })
    await check('2', 'S3 largeur Photo 6 cm (aller-retour)', async () => {
      const r = ws.getRange('C:C')
      r.format.columnWidth = cm(6)
      await ctx.sync()
      r.format.load('columnWidth')
      await ctx.sync()
      return `demandé ${cm(6)}, relu ${r.format.columnWidth}`
    })
    await check('2', 'S4 hauteur de ligne 130 pt (aller-retour)', async () => {
      const r = ws.getRange('2:2')
      r.format.rowHeight = 130
      await ctx.sync()
      r.format.load('rowHeight')
      await ctx.sync()
      return `relu ${r.format.rowHeight}`
    })
    await check('2', 'S5 position de C3 (Range.left/top)', async () => {
      const r = ws.getRange('C3')
      r.load('left,top,width,height')
      await ctx.sync()
      return `left ${r.left}, top ${r.top}, width ${r.width}, height ${r.height}`
    })
    await check('2', 'S6 format date dd/mm/yyyy donné en une seule valeur', async () => {
      const r = t.columns.getItem('Échéance').getDataBodyRange()
      r.numberFormat = 'dd/mm/yyyy'
      await ctx.sync()
      const c = ws.getRange('F2')
      c.load('text,values')
      await ctx.sync()
      return `F2 affiche « ${c.text[0][0]} » (attendu 05/10/2026), valeur ${c.values[0][0]}`
    })
    await check('2', 'S6b format texte @ : « =1+1 » et « - joint » restent du texte', async () => {
      const r = ws.getRange('D2:D3')
      r.numberFormat = '@'
      r.values = [['=1+1'], ['- joint']]
      await ctx.sync()
      r.load('values,formulas')
      await ctx.sync()
      return `valeurs ${JSON.stringify(r.values)}, formules ${JSON.stringify(r.formulas)} (attendu : texte identique, pas 2)`
    })
    await check('2', 'retour à la ligne, police 9, alignement haut', async () => {
      const r = t.getRange()
      r.format.wrapText = true
      r.format.font.size = 9
      r.format.verticalAlignment = 'Top'
      await ctx.sync()
    })
    await check('2', 'en-tête sombre, première ligne figée', async () => {
      const h = t.getHeaderRowRange()
      h.format.fill.color = '#2B2B2B'
      h.format.font.color = '#FFFFFF'
      h.format.font.bold = true
      ws.freezePanes.freezeRows(1)
      await ctx.sync()
    })
    await check('2', 'S7 liste déroulante Pièce acceptant une autre valeur', async () => {
      const lists = ctx.workbook.worksheets.add(LISTS)
      lists.getRange('A1:A4').values = [['Pièce'], ['Cuisine'], ['Entrée'], ['Séjour']]
      const r = ws.getRange('E2:E1048576')
      r.dataValidation.clear()
      r.dataValidation.rule = { list: { inCellDropDown: true, source: `=${LISTS}!$A$2:$A$4` } }
      r.dataValidation.errorAlert = { showAlert: false, style: 'Information', title: 'Action !', message: 'Valeur ajoutée.' }
      await ctx.sync()
      return 'tapez « Garage » dans E5 : la valeur doit être acceptée sans message'
    })
    await check('2', 'S8 couleur de statut (mise en forme conditionnelle)', async () => {
      const cf = ws.getRange('G2:G1048576').conditionalFormats.add(Excel.ConditionalFormatType.cellValue)
      cf.cellValue.format.fill.color = '#34A853'
      cf.cellValue.format.font.color = '#FFFFFF'
      cf.cellValue.rule = { formula1: '="Fait"', operator: 'EqualTo' }
      await ctx.sync()
    })
    await check('2', 'S9 trait + pièce en gras au changement de pièce', async () => {
      const line = ws.getRange('B2:G1048576').conditionalFormats.add(Excel.ConditionalFormatType.custom)
      line.custom.rule.formula = '=AND($E2<>"",$E2<>$E1)'
      const top = line.custom.format.borders.getItem('EdgeTop')
      top.style = 'Continuous'
      top.color = '#404040'
      await ctx.sync()
      const bold = ws.getRange('E2:E1048576').conditionalFormats.add(Excel.ConditionalFormatType.custom)
      bold.custom.rule.formula = '=AND($E2<>"",$E2<>$E1)'
      bold.custom.format.font.bold = true
      await ctx.sync()
      return 'triez par Pièce : un trait fin doit séparer les pièces, la pièce en gras'
    })
    await check('2', 'S10 mise en page A4 portrait, marges, en-tête et pied', async () => {
      const p = ws.pageLayout
      p.orientation = Excel.PageOrientation.portrait
      p.paperSize = Excel.PaperType.a4
      p.leftMargin = cm(1)
      p.rightMargin = cm(1)
      p.topMargin = cm(1.9)
      p.bottomMargin = cm(1.5)
      p.headerMargin = cm(0.8)
      p.footerMargin = cm(0.6)
      p.centerHorizontally = true
      p.headersFooters.defaultForAllPages.leftHeader = 'Maison Dupont && fils'
      p.headersFooters.defaultForAllPages.rightHeader = '&D'
      p.headersFooters.defaultForAllPages.centerFooter = 'Page &P/&N'
      p.setPrintTitleRows('$1:$1')
      p.setPrintArea('$B:$G')
      await ctx.sync()
      return 'Fichier › Imprimer : A4 portrait, « Maison Dupont & fils », date, « Page 1/N », colonne ID absente'
    })
    await check('2', 'S11 ajuster à 1 page en largeur', async () => {
      ws.pageLayout.zoom = { horizontalFitToPages: 1, verticalFitToPages: 0 }
      await ctx.sync()
      ws.pageLayout.load('zoom')
      await ctx.sync()
      return JSON.stringify(ws.pageLayout.zoom)
    })
    await check('2', 'S12 feuille très masquée', async () => {
      const found = ctx.workbook.worksheets.getItemOrNullObject('_essai')
      await ctx.sync()
      const hidden = found.isNullObject ? ctx.workbook.worksheets.add('_essai') : found
      hidden.getRange('A1').values = [['J{"essai":true}']]
      hidden.visibility = Excel.SheetVisibility.veryHidden
      await ctx.sync()
      return 'clic droit sur un onglet › Afficher : « _essai » ne doit pas apparaître'
    })
  }),
)

on('b3', () =>
  Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(SHEET)
    const column = ws.getRange('C:C')
    column.load('left,width')
    for (let i = 0; i < 3; i++) ws.getRange(`${i + 2}:${i + 2}`).format.rowHeight = 130
    await ctx.sync()
    for (let i = 0; i < 3; i++) {
      const cell = ws.getRange(`C${i + 2}`)
      cell.load('top')
      await ctx.sync()
      const s = ws.shapes.addImage(testJpeg(472, 354, `id-${i + 1}`))
      s.name = `img_test${i + 1}`
      s.lockAspectRatio = false
      s.left = column.left + 3
      s.top = cell.top + 3
      s.width = column.width - 6
      s.height = (column.width - 6) * 0.75
      s.placement = Excel.Placement.twoCell
      await ctx.sync()
      s.load('name,left,top,width,height,placement')
      await ctx.sync()
      log('3', `S13 ${s.name} : left ${s.left}, top ${s.top}, ${s.width}×${s.height}, ancrage ${s.placement} (ligne ${i + 2} top ${cell.top})`)
    }
    const found = ws.shapes.getItemOrNullObject('img_test2')
    await ctx.sync()
    log('3', `S13 recherche par nom img_test2 : ${found.isNullObject ? 'introuvable' : 'trouvée'}`)
  }),
)

on('b4', () =>
  Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(SHEET)
    const t = ctx.workbook.tables.getItem(TABLE)
    const describe = (kind) => async (e) => {
      eventCount++
      const details = e.details ? ` avant=${JSON.stringify(e.details.valueBefore)} après=${JSON.stringify(e.details.valueAfter)}` : ''
      log('4', `S14 ${kind} #${eventCount} ${e.type} ${e.address || ''} ${e.changeType || ''} source=${e.source || ''} déclencheur=${e.triggerSource || ''}${details}`)
    }
    const registrations = [
      ['modification du tableau', () => t.onChanged.add(describe('tableau'))],
      ['modification de la feuille', () => ws.onChanged.add(describe('feuille'))],
      ['modification du classeur', () => ctx.workbook.worksheets.onChanged.add(describe('classeur'))],
      ['tri des lignes', () => ctx.workbook.worksheets.onRowSorted.add(describe('tri'))],
      ['filtre', () => ws.onFiltered.add(describe('filtre'))],
      ['ligne masquée', () => ws.onRowHiddenChanged.add(describe('masquage'))],
    ]
    for (const [label, register] of registrations) {
      await check('4', `abonnement ${label}`, async () => {
        register()
        await ctx.sync()
      })
    }
  }),
)

on('b4b', () =>
  Excel.run(async (ctx) => {
    const before = eventCount
    ctx.runtime.enableEvents = false
    ctx.workbook.worksheets.getItem(SHEET).getRange('D4').values = [['Écrit par le complément']]
    await ctx.sync()
    ctx.runtime.enableEvents = true
    await ctx.sync()
    setTimeout(() => log('4b', `S15 événements reçus pendant l’écriture du complément : ${eventCount - before} (attendu 0)`), 1500)
  }),
)

on('b5', () =>
  Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(SHEET)
    const body = ctx.workbook.tables.getItem(TABLE).getDataBodyRange()
    body.load('values,rowCount')
    ws.shapes.load('items/name,items/top,items/height')
    await ctx.sync()
    const rows = Array.from({ length: body.rowCount }, (_, i) => body.getRow(i).load('top,height'))
    await ctx.sync()
    for (const s of ws.shapes.items.filter((x) => x.name.startsWith('img_test'))) {
      const i = rows.findIndex((r) => s.top + 3 >= r.top && s.top + 3 < r.top + Math.max(r.height, 0.1))
      log('5', `S16 ${s.name} (posée pour ${s.name.replace('img_test', 'id-')}) est sur la ligne de ${i >= 0 ? body.values[i][0] : '?'} ; hauteur image ${s.height}`)
    }
    log('5', 'S16 notez aussi : avec le filtre « Pièce = Cuisine », les photos des lignes masquées ont-elles disparu (OUI/NON) ?')
  }),
)

on('b5b', () =>
  Excel.run(async (ctx) => {
    const body = ctx.workbook.tables.getItem(TABLE).getDataBodyRange()
    body.load('rowCount')
    await ctx.sync()
    const rows = Array.from({ length: body.rowCount }, (_, i) => body.getRow(i).load('height'))
    await ctx.sync()
    const hiddenBefore = rows.filter((r) => r.height === 0).length
    body.format.autofitRows()
    await ctx.sync()
    rows.forEach((r) => r.load('height'))
    await ctx.sync()
    const hiddenAfter = rows.filter((r) => r.height === 0).length
    log('5b', `S17 lignes masquées par le filtre : avant ${hiddenBefore}, après ajustement ${hiddenAfter} (attendu identiques)`)
  }),
)

async function insertMany(total, batch) {
  return Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(SHEET)
    const image = testJpeg(630, 472, 'perf')
    const prefix = `img_perf${Date.now()}_`
    const start = performance.now()
    for (let done = 0; done < total; done += batch) {
      for (let i = done; i < Math.min(total, done + batch); i++) {
        const s = ws.shapes.addImage(image)
        s.name = `${prefix}${i}`
        s.left = 600 + (i % 10) * 20
        s.top = 20 + Math.floor(i / 10) * 20
        s.width = 80
        s.height = 60
      }
      await ctx.sync()
    }
    return Math.round(performance.now() - start)
  })
}
on('b6', async () => log('6', `S18 40 photos (630 px) en un lot : ${await insertMany(40, 40)} ms`))
on('b6b', async () => log('6b', `S18 300 photos par lots de 40 : ${await insertMany(300, 40)} ms — enregistrez le classeur et notez sa taille`))

on('b7', async () => {
  const stamp = new Date().toLocaleString('fr-FR')
  localStorage.setItem('action-spike', stamp)
  await idbPut(stamp)
  log('7', `S19 écrit maintenant : ${stamp}. Fermez Excel complètement, rouvrez ce classeur et le volet : les lignes [7] du haut doivent afficher cette valeur.`)
})

on('b8', async () => {
  if (!window.supabase) throw new Error('bibliothèque Supabase non chargée (réseau ?)')
  const sb = client()
  const email = document.getElementById('email').value.trim()
  const password = document.getElementById('password').value
  if (email && password) {
    const { data, error } = await sb.auth.signInWithPassword({ email, password })
    if (error) throw error
    log('8', `S20 connecté : ${data.user.email}`)
  }
  const { data: current } = await sb.auth.getSession()
  if (!current.session) throw new Error('aucune session : saisissez e-mail et mot de passe')
  const { data, error } = await sb.rpc('pull_changes', { p_since: 0, p_limit: 1000 })
  if (error) throw error
  const items = data.items || []
  log('8', `S20 ${items.length} élément(s) reçus du serveur`)
  const photo = items.find((i) => i.kind === 'photo' && (i.row.data.renderedPath || i.row.data.originalPath))
  if (!photo) {
    log('8', 'S20 aucune photo sur le serveur : prenez une photo avec l’iPhone puis recommencez (sinon notez « non testé »)')
    return
  }
  const path = photo.row.data.renderedPath || photo.row.data.originalPath
  const t0 = performance.now()
  const download = await sb.storage.from('photos').download(path)
  if (download.error) throw download.error
  const bitmap = await createImageBitmap(download.data)
  const w = 472
  const h = Math.round((bitmap.height * w) / bitmap.width)
  const canvas = document.createElement('canvas')
  canvas.width = w
  canvas.height = h
  canvas.getContext('2d').drawImage(bitmap, 0, 0, w, h)
  const url = canvas.toDataURL('image/jpeg', 0.8)
  log('8', `S20 photo ${path} téléchargée et réduite à ${w}×${h} (${Math.round(url.length / 1024)} Ko en base64) en ${Math.round(performance.now() - t0)} ms`)
  await Excel.run(async (ctx) => {
    const s = ctx.workbook.worksheets.getItem(SHEET).shapes.addImage(url.slice(url.indexOf(',') + 1))
    s.name = `img_reel${Date.now()}`
    s.left = 300
    s.top = 20
    await ctx.sync()
  })
  log('8', 'S20 photo réelle insérée dans la feuille Essai')
})

on('b9', async () => {
  await Excel.run(async (ctx) => {
    ctx.workbook.settings.add('action.essai', new Date().toLocaleString('fr-FR'))
    await ctx.sync()
  })
  Office.context.document.settings.set('Office.AutoShowTaskpaneWithDocument', true)
  await new Promise((resolve, reject) =>
    Office.context.document.settings.saveAsync((r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(r.error))),
  )
  log('9', 'S21 réglage écrit et ouverture automatique demandée : enregistrez (Ctrl+S), fermez le classeur, rouvrez-le : le volet doit s’ouvrir seul et la ligne [9] du haut afficher la date.')
})

on('copy', async () => logArea.select())

import { useMemo, useRef, useState } from 'react'
import * as XLSX from 'xlsx'
import { jsPDF } from 'jspdf'
import * as pdfjsLib from 'pdfjs-dist'
import pdfjsWorker from 'pdfjs-dist/build/pdf.worker.min.mjs?url'
import './App.css'

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorker

const directCostDefaults = [
  {
    key: 'blank',
    name: 'Blank garment/product',
    unitCost: 3.5,
    qty: 24,
    include: true,
  },
  {
    key: 'transfer',
    name: 'Transfer / ink / vinyl / thread',
    unitCost: 1.25,
    qty: 24,
    include: true,
  },
  { key: 'packaging', name: 'Packaging', unitCost: 0.35, qty: 24, include: true },
  { key: 'outside', name: 'Outside services', unitCost: 0, qty: 1, include: true },
  { key: 'shipping', name: 'Shipping / freight', unitCost: 0, qty: 1, include: true },
  { key: 'setup', name: 'Setup / artwork', unitCost: 15, qty: 1, include: true },
  {
    key: 'other',
    name: 'Other direct cost',
    unitCost: 0,
    qty: 1,
    include: true,
  },
]

const laborDefaults = [
  {
    key: 'design',
    name: 'Design / prepress labor',
    rateOrPercent: 25,
    baseHours: 0.5,
    include: true,
  },
  {
    key: 'production',
    name: 'Production labor',
    rateOrPercent: 22,
    baseHours: 1.5,
    include: true,
  },
  {
    key: 'finishing',
    name: 'Finishing / packing labor',
    rateOrPercent: 18,
    baseHours: 0.5,
    include: true,
  },
  {
    key: 'machine',
    name: 'Machine / production overhead',
    rateOrPercent: 12,
    baseHours: 1.5,
    include: true,
  },
  {
    key: 'overhead',
    name: 'General overhead allocation %',
    rateOrPercent: 10,
    baseHours: 0,
    include: true,
  },
  {
    key: 'waste',
    name: 'Waste / spoilage allowance %',
    rateOrPercent: 5,
    baseHours: 0,
    include: true,
  },
  {
    key: 'payment',
    name: 'Payment processing %',
    rateOrPercent: 3,
    baseHours: 0,
    include: true,
  },
  {
    key: 'rush',
    name: 'Rush / special handling %',
    rateOrPercent: 0,
    baseHours: 0,
    include: false,
  },
]

const productionMethods = [
  'DTF / Heat Transfer',
  'HTV / Vinyl',
  'Screen Printing',
  'Embroidery',
  'Sublimation',
  'Direct-to-Garment',
]

const assumptionsDefault = {
  hourlyProductionLaborRate: 20,
  laborBurdenPercent: 18,
  monthlyFixedOverhead: 6500,
  monthlyProductiveShopHours: 520,
  targetGrossMarginPercent: 50,
  defaultWastePercent: 3,
  rushFeePercent: 20,
  salesFeePercent: 3.5,
  setupArtworkRate: 45,
  qcRate: 22,
  machineRates: {
    'DTF / Heat Transfer': 12,
    'HTV / Vinyl': 8,
    'Screen Printing': 18,
    Embroidery: 22,
    Sublimation: 14,
    'Direct-to-Garment': 28,
  },
}

function asNumber(value) {
  const n = Number(value)
  return Number.isFinite(n) ? n : 0
}

function normalizeNumericInput(value) {
  return value === '' ? '' : asNumber(value)
}

function normalizeKey(value) {
  return String(value ?? '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
}

function asBoolean(value) {
  if (typeof value === 'boolean') return value
  if (typeof value === 'string') return ['true', '1', 'yes', 'y'].includes(value.trim().toLowerCase())
  return Boolean(value)
}

function money(value) {
  return new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: 'USD',
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(value)
}

function pct(value) {
  return `${value.toFixed(1)}%`
}

function buildEmptyJobRow(id) {
  return {
    id,
    customer: '',
    method: 'DTF / Heat Transfer',
    qty: 0,
    blankCost: 0,
    printMaterialCost: 0,
    setupHours: 0,
    productionHours: 0,
    qcPackHours: 0,
    shipping: 0,
    otherDirect: 0,
    rush: false,
    quotedPrice: 0,
  }
}

function calculateCostCalculator(calcState) {
  const quantity = Math.max(0, asNumber(calcState.quantity))
  const targetMarginDecimal = asNumber(calcState.targetGrossMarginPercent) / 100

  const directRows = calcState.directRows.map((row) => {
    const extended = row.include ? asNumber(row.unitCost) * asNumber(row.qty) : 0
    return { ...row, extended }
  })

  const laborRows = calcState.laborRows.map((row, idx) => {
    if (idx <= 3) {
      return {
        ...row,
        calculated: row.include
          ? asNumber(row.rateOrPercent) * asNumber(row.baseHours)
          : 0,
      }
    }

    if (idx === 4) {
      const subtotal =
        directRows.reduce((sum, row) => sum + row.extended, 0) +
        laborRowsSafeBase(calcState.laborRows)
      return {
        ...row,
        calculated: row.include ? subtotal * (asNumber(row.rateOrPercent) / 100) : 0,
      }
    }

    if (idx === 5) {
      const blankAndMaterial = (directRows[0]?.extended || 0) + (directRows[1]?.extended || 0)
      return {
        ...row,
        calculated: row.include
          ? blankAndMaterial * (asNumber(row.rateOrPercent) / 100)
          : 0,
      }
    }

    return { ...row, calculated: 0 }
  })

  const totalTrueJobCost =
    directRows.reduce((sum, row) => sum + row.extended, 0) +
    laborRows.reduce((sum, row, idx) => sum + (idx <= 5 ? row.calculated : 0), 0)

  const trueCostPerItem = quantity > 0 ? totalTrueJobCost / quantity : 0

  const paymentProcessingPercent =
    laborRows[6]?.include ? asNumber(laborRows[6]?.rateOrPercent) / 100 : 0

  const breakEvenPricePerItem = trueCostPerItem
  const targetDenominator = 1 - targetMarginDecimal - paymentProcessingPercent
  const targetSellingPricePerItem =
    targetDenominator > 0 ? trueCostPerItem / targetDenominator : 0

  const rushPercent = laborRows[7]?.include ? asNumber(laborRows[7]?.rateOrPercent) / 100 : 0
  const recommendedJobTotal = targetSellingPricePerItem * quantity * (1 + rushPercent)

  const expectedGrossProfit =
    recommendedJobTotal - totalTrueJobCost - recommendedJobTotal * paymentProcessingPercent
  const expectedGrossMargin = recommendedJobTotal > 0 ? expectedGrossProfit / recommendedJobTotal : 0

  const testPricePerItem = asNumber(calcState.testPricePerItem)
  const quotedJobTotal = testPricePerItem * quantity
  const paymentFee = quotedJobTotal * paymentProcessingPercent
  const profitAfterCosts = quotedJobTotal - totalTrueJobCost - paymentFee
  const profitPerItem = quantity > 0 ? profitAfterCosts / quantity : 0
  const grossMargin = quotedJobTotal > 0 ? profitAfterCosts / quotedJobTotal : 0
  const differenceVsTarget = testPricePerItem - targetSellingPricePerItem

  let status = 'Enter price'
  if (testPricePerItem > 0) {
    status = testPricePerItem >= targetSellingPricePerItem ? 'Meets target' : 'Below target'
  }

  return {
    directRows,
    laborRows,
    totalTrueJobCost,
    trueCostPerItem,
    targetGrossMarginPercent: targetMarginDecimal * 100,
    breakEvenPricePerItem,
    targetSellingPricePerItem,
    recommendedJobTotal,
    expectedGrossProfit,
    expectedGrossMargin,
    testPricePerItem,
    quotedJobTotal,
    paymentFee,
    profitAfterCosts,
    profitPerItem,
    grossMargin,
    differenceVsTarget,
    status,
  }
}

function laborRowsSafeBase(laborRows) {
  return laborRows.slice(0, 4).reduce((sum, row) => {
    const rowCost = row.include ? asNumber(row.rateOrPercent) * asNumber(row.baseHours) : 0
    return sum + rowCost
  }, 0)
}

function roundUp(value) {
  return Math.ceil(value)
}

function calculateJobRow(row, assumptions) {
  const qty = Math.max(0, asNumber(row.qty))
  const blankCost = asNumber(row.blankCost)
  const printMaterialCost = asNumber(row.printMaterialCost)
  const setupHours = asNumber(row.setupHours)
  const productionHours = asNumber(row.productionHours)
  const qcPackHours = asNumber(row.qcPackHours)
  const shipping = asNumber(row.shipping)
  const otherDirect = asNumber(row.otherDirect)
  const quotedPrice = asNumber(row.quotedPrice)

  const wasteRate = asNumber(assumptions.defaultWastePercent) / 100
  const targetMargin = asNumber(assumptions.targetGrossMarginPercent) / 100
  const rushFee = asNumber(assumptions.rushFeePercent) / 100
  const salesFee = asNumber(assumptions.salesFeePercent) / 100
  const setupRate = asNumber(assumptions.setupArtworkRate)
  const qcRate = asNumber(assumptions.qcRate)

  const loadedLaborRate =
    asNumber(assumptions.hourlyProductionLaborRate) *
      (1 + asNumber(assumptions.laborBurdenPercent) / 100) +
    (asNumber(assumptions.monthlyProductiveShopHours) > 0
      ? asNumber(assumptions.monthlyFixedOverhead) /
        asNumber(assumptions.monthlyProductiveShopHours)
      : 0)

  const machineRate = asNumber(assumptions.machineRates[row.method] || 0)

  const wasteQty = roundUp(qty * wasteRate)
  const blankAndMaterialCost = (qty + wasteQty) * (blankCost + printMaterialCost)
  const setupLabor = setupHours * setupRate
  const productionLaborAndOverhead = productionHours * loadedLaborRate
  const machineCost = productionHours * machineRate
  const qcPackLabor = qcPackHours * qcRate
  const salesFees = quotedPrice * salesFee
  const rushSurcharge = row.rush ? quotedPrice * rushFee : 0

  const totalJobCost =
    blankAndMaterialCost +
    setupLabor +
    productionLaborAndOverhead +
    machineCost +
    qcPackLabor +
    shipping +
    otherDirect +
    salesFees

  const costPerGoodUnit = qty > 0 ? totalJobCost / qty : 0
  const recommendedPrice =
    qty > 0
      ? 1 - targetMargin - salesFee > 0
        ? ((totalJobCost - salesFees) / (1 - targetMargin - salesFee)) *
          (row.rush ? 1 + rushFee : 1)
        : 0
      : 0

  const actualGrossProfit = quotedPrice - totalJobCost
  const actualMarginPercent = quotedPrice > 0 ? actualGrossProfit / quotedPrice : 0

  return {
    wasteQty,
    blankAndMaterialCost,
    setupLabor,
    productionLaborAndOverhead,
    machineCost,
    qcPackLabor,
    salesFees,
    rushSurcharge,
    totalJobCost,
    costPerGoodUnit,
    recommendedPrice,
    actualGrossProfit,
    actualMarginPercent,
    loadedLaborRate,
  }
}

function App() {
  const [activeTab, setActiveTab] = useState('cost')
  const [guideFocus, setGuideFocus] = useState('cost')
  const [importModalOpen, setImportModalOpen] = useState(false)
  const [exportModalOpen, setExportModalOpen] = useState(false)
  const [templateModalOpen, setTemplateModalOpen] = useState(false)
  const [templateSide, setTemplateSide] = useState('cost')
  const [pendingImportSide, setPendingImportSide] = useState(null)
  const fileInputRef = useRef(null)

  const [costState, setCostState] = useState({
    jobCustomer: '',
    product: 'T-Shirt',
    productionMethod: 'DTF',
    quantity: 24,
    sidesPrinted: 1,
    colorComplexity: 1,
    directRows: directCostDefaults,
    laborRows: laborDefaults,
    targetGrossMarginPercent: 40,
    testPricePerItem: 0,
  })

  const [assumptions, setAssumptions] = useState(assumptionsDefault)

  const [jobRows, setJobRows] = useState(
    Array.from({ length: 6 }, (_, i) => buildEmptyJobRow(`job-${i + 1}`)),
  )

  const [quickQuote, setQuickQuote] = useState({
    quantity: 24,
    blankCostPerUnit: 5.5,
    decorationCostPerUnit: 2.25,
    setupArtwork: 35,
    productionOverhead: 60,
    shippingOther: 15,
    targetMarginPercent: 50,
  })

  const costResults = useMemo(() => calculateCostCalculator(costState), [costState])

  const calculatedRows = useMemo(
    () => jobRows.map((row) => ({ ...row, result: calculateJobRow(row, assumptions) })),
    [jobRows, assumptions],
  )

  const quickQuoteResults = useMemo(() => {
    const qty = Math.max(0, asNumber(quickQuote.quantity))
    const totalCost =
      qty *
        (asNumber(quickQuote.blankCostPerUnit) +
          asNumber(quickQuote.decorationCostPerUnit)) +
      asNumber(quickQuote.setupArtwork) +
      asNumber(quickQuote.productionOverhead) +
      asNumber(quickQuote.shippingOther)

    const costPerUnit = qty > 0 ? totalCost / qty : 0
    const margin = asNumber(quickQuote.targetMarginPercent) / 100
    const recommendedSellingPrice = 1 - margin > 0 ? totalCost / (1 - margin) : 0
    const recommendedPricePerUnit = qty > 0 ? recommendedSellingPrice / qty : 0

    return {
      totalCost,
      costPerUnit,
      recommendedSellingPrice,
      recommendedPricePerUnit,
    }
  }, [quickQuote])

  function updateCostField(field, value) {
    setCostState((prev) => ({ ...prev, [field]: normalizeNumericInput(value) }))
  }

  function updateDirectRow(index, field, value) {
    setCostState((prev) => {
      const nextRows = [...prev.directRows]
      nextRows[index] = { ...nextRows[index], [field]: normalizeNumericInput(value) }
      return { ...prev, directRows: nextRows }
    })
  }

  function updateLaborRow(index, field, value) {
    setCostState((prev) => {
      const nextRows = [...prev.laborRows]
      nextRows[index] = { ...nextRows[index], [field]: normalizeNumericInput(value) }
      return { ...prev, laborRows: nextRows }
    })
  }

  function updateAssumption(field, value) {
    setAssumptions((prev) => ({ ...prev, [field]: normalizeNumericInput(value) }))
  }

  function updateMachineRate(method, value) {
    setAssumptions((prev) => ({
      ...prev,
      machineRates: {
        ...prev.machineRates,
        [method]: normalizeNumericInput(value),
      },
    }))
  }

  function updateJobRow(id, field, value) {
    setJobRows((prev) =>
      prev.map((row) => {
        const nextValue = field === 'customer' || field === 'method' ? value : normalizeNumericInput(value)
        return row.id === id ? { ...row, [field]: nextValue } : row
      }),
    )
  }

  function addJobRow() {
    setJobRows((prev) => [...prev, buildEmptyJobRow(`job-${prev.length + 1}`)])
  }

  function updateQuickQuote(field, value) {
    setQuickQuote((prev) => ({ ...prev, [field]: normalizeNumericInput(value) }))
  }

  function openUsageGuide(target) {
    setGuideFocus(target)
    setActiveTab('guide')
  }

  function triggerImport(side) {
    setPendingImportSide(side)
    setImportModalOpen(false)
    setTimeout(() => fileInputRef.current?.click(), 0)
  }

  async function extractPdfText(file) {
    const buffer = await file.arrayBuffer()
    const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(buffer) }).promise
    let text = ''

    for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
      const page = await pdf.getPage(pageNumber)
      const content = await page.getTextContent()
      text += content.items.map((item) => item.str).join(' ') + ' '
    }

    return text
  }

  function parsePdfImportText(text, side) {
    const upper = text.toLowerCase()
    const getValue = (patterns) => {
      for (const pattern of patterns) {
        const match = upper.match(pattern)
        if (match) return match[1]
      }
      return null
    }

    if (side === 'cost') {
      const settings = {
        jobCustomer: getValue([/job\s*\/\s*customer\s*[:\-]?\s*([\w\s&]+)/i]),
        product: getValue([/product\s*\/\s*garment\s*[:\-]?\s*([\w\s&]+)/i]),
        productionMethod: getValue([/production\s*method\s*[:\-]?\s*([\w\s&\/]+)/i]),
        quantity: Number.parseFloat(getValue([/quantity\s*[:\-]?\s*(\d+(?:\.\d+)?)/i]) || '0'),
        sidesPrinted: Number.parseFloat(getValue([/sides\s*printed\s*[:\-]?\s*(\d+(?:\.\d+)?)/i]) || '0'),
        colorComplexity: Number.parseFloat(getValue([/colors\s*\/\s*complexity\s*[:\-]?\s*(\d+(?:\.\d+)?)/i]) || '0'),
        targetGrossMarginPercent: Number.parseFloat(getValue([/target\s*gross\s*margin\s*[:\-]?\s*(\d+(?:\.\d+)?)/i]) || '0'),
        testPricePerItem: Number.parseFloat(getValue([/your\s*price\s*\/\s*item\s*[:\-]?\s*(\d+(?:\.\d+)?)/i]) || '0'),
      }

      const directRows = [
        { key: 'blank', name: 'Blank garment/product', unitCost: 0, qty: 0, include: true },
        { key: 'transfer', name: 'Transfer / ink / vinyl / thread', unitCost: 0, qty: 0, include: true },
        { key: 'packaging', name: 'Packaging', unitCost: 0, qty: 0, include: true },
        { key: 'outside', name: 'Outside services', unitCost: 0, qty: 0, include: true },
        { key: 'shipping', name: 'Shipping / freight', unitCost: 0, qty: 0, include: true },
        { key: 'setup', name: 'Setup / artwork', unitCost: 0, qty: 0, include: true },
        { key: 'other', name: 'Other direct cost', unitCost: 0, qty: 0, include: true },
      ]

      const laborRows = [
        { key: 'design', name: 'Design / prepress labor', rateOrPercent: 0, baseHours: 0, include: true },
        { key: 'production', name: 'Production labor', rateOrPercent: 0, baseHours: 0, include: true },
        { key: 'finishing', name: 'Finishing / packing labor', rateOrPercent: 0, baseHours: 0, include: true },
        { key: 'machine', name: 'Machine / production overhead', rateOrPercent: 0, baseHours: 0, include: true },
        { key: 'overhead', name: 'General overhead allocation %', rateOrPercent: 0, baseHours: 0, include: true },
        { key: 'waste', name: 'Waste / spoilage allowance %', rateOrPercent: 0, baseHours: 0, include: true },
        { key: 'payment', name: 'Payment processing %', rateOrPercent: 0, baseHours: 0, include: true },
        { key: 'rush', name: 'Rush / special handling %', rateOrPercent: 0, baseHours: 0, include: false },
      ]

      return { directRows, laborRows, costState: settings }
    }

    const settings = {
      hourlyProductionLaborRate: Number.parseFloat(getValue([/hourly\s*production\s*labor\s*rate\s*[:\-]?\s*(\d+(?:\.\d+)?)/i]) || '0'),
      laborBurdenPercent: Number.parseFloat(getValue([/labor\s*burden\s*[:\-]?\s*(\d+(?:\.\d+)?)/i]) || '0'),
      monthlyFixedOverhead: Number.parseFloat(getValue([/monthly\s*fixed\s*overhead\s*[:\-]?\s*(\d+(?:\.\d+)?)/i]) || '0'),
      monthlyProductiveShopHours: Number.parseFloat(getValue([/monthly\s*productive\s*shop\s*hours\s*[:\-]?\s*(\d+(?:\.\d+)?)/i]) || '0'),
      targetGrossMarginPercent: Number.parseFloat(getValue([/target\s*gross\s*margin\s*[:\-]?\s*(\d+(?:\.\d+)?)/i]) || '0'),
      defaultWastePercent: Number.parseFloat(getValue([/default\s*waste\s*[:\-]?\s*(\d+(?:\.\d+)?)/i]) || '0'),
      rushFeePercent: Number.parseFloat(getValue([/rush\s*fee\s*[:\-]?\s*(\d+(?:\.\d+)?)/i]) || '0'),
      salesFeePercent: Number.parseFloat(getValue([/sales\s*\/\s*payment\s*processing\s*[:\-]?\s*(\d+(?:\.\d+)?)/i]) || '0'),
      setupArtworkRate: Number.parseFloat(getValue([/setup\s*\/\s*artwork\s*rate\s*[:\-]?\s*(\d+(?:\.\d+)?)/i]) || '0'),
      qcRate: Number.parseFloat(getValue([/qc\s*rate\s*[:\-]?\s*(\d+(?:\.\d+)?)/i]) || '0'),
    }

    return {
      assumptions: settings,
      jobRows: [],
      quickQuote: {
        quantity: 0,
        blankCostPerUnit: 0,
        decorationCostPerUnit: 0,
        setupArtwork: 0,
        productionOverhead: 0,
        shippingOther: 0,
        targetMarginPercent: 0,
      },
    }
  }

  function getSheetByName(workbook, names) {
    const lookupNames = names.map((name) => name.toLowerCase())
    const target = workbook.SheetNames.find((sheetName) => lookupNames.includes(sheetName.toLowerCase()))
    return target ? workbook.Sheets[target] : null
  }

  function buildAExportWorkbook() {
    const workbook = XLSX.utils.book_new()
    const results = costResults

    const settings = [
      { Label: 'Job / Customer', Value: costState.jobCustomer },
      { Label: 'Product / Garment', Value: costState.product },
      { Label: 'Production Method', Value: costState.productionMethod },
      { Label: 'Quantity', Value: costState.quantity },
      { Label: 'Sides Printed', Value: costState.sidesPrinted },
      { Label: 'Colors / Complexity', Value: costState.colorComplexity },
      { Label: 'Target Gross Margin %', Value: costState.targetGrossMarginPercent },
      { Label: 'Your Price / Item (test)', Value: costState.testPricePerItem },
    ]

    const directRows = costState.directRows.map((row, index) => ({
      'Cost Component': row.name,
      'Unit Cost': row.unitCost,
      Qty: row.qty,
      Include: row.include ? 'Yes' : 'No',
      'Extended Cost': results.directRows[index]?.extended || 0,
    }))

    const laborRows = costState.laborRows.map((row, index) => ({
      'Cost Driver': row.name,
      'Rate / %': row.rateOrPercent,
      'Hours / Base': row.baseHours,
      Include: row.include ? 'Yes' : 'No',
      'Calculated Cost': results.laborRows[index]?.calculated || 0,
    }))

    const resultRows = [
      { Label: 'Total True Job Cost', Value: results.totalTrueJobCost },
      { Label: 'True Cost Per Item', Value: results.trueCostPerItem },
      { Label: 'Break-even Price / Item', Value: results.breakEvenPricePerItem },
      { Label: 'Target Selling Price / Item', Value: results.targetSellingPricePerItem },
      { Label: 'Recommended Job Total', Value: results.recommendedJobTotal },
      { Label: 'Expected Gross Profit', Value: results.expectedGrossProfit },
      { Label: 'Expected Gross Margin', Value: results.expectedGrossMargin },
      { Label: 'Quoted Job Total', Value: results.quotedJobTotal },
      { Label: 'Payment Fee', Value: results.paymentFee },
      { Label: 'Profit After Costs', Value: results.profitAfterCosts },
      { Label: 'Profit / Item', Value: results.profitPerItem },
      { Label: 'Test Gross Margin', Value: results.grossMargin },
      { Label: 'Difference vs Target', Value: results.differenceVsTarget },
      { Label: 'Status', Value: results.status },
    ]

    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(settings), 'Settings')
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(directRows), 'DirectCosts')
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(laborRows), 'LaborCosts')
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(resultRows), 'Results')

    return workbook
  }

  function buildBExportWorkbook() {
    const workbook = XLSX.utils.book_new()

    const assumptionsSheet = [
      { Label: 'Hourly Production Labor Rate', Value: assumptions.hourlyProductionLaborRate },
      { Label: 'Labor Burden %', Value: assumptions.laborBurdenPercent },
      { Label: 'Monthly Fixed Overhead', Value: assumptions.monthlyFixedOverhead },
      { Label: 'Monthly Productive Shop Hours', Value: assumptions.monthlyProductiveShopHours },
      { Label: 'Target Gross Margin %', Value: assumptions.targetGrossMarginPercent },
      { Label: 'Default Waste %', Value: assumptions.defaultWastePercent },
      { Label: 'Rush Fee %', Value: assumptions.rushFeePercent },
      { Label: 'Sales / Payment Processing %', Value: assumptions.salesFeePercent },
      { Label: 'Setup / Artwork Rate per Hour', Value: assumptions.setupArtworkRate },
      { Label: 'QC Rate per Hour', Value: assumptions.qcRate },
    ]

    const machineRates = Object.entries(assumptions.machineRates).map(([method, rate]) => ({
      Method: method,
      Rate: rate,
    }))

    const jobs = jobRows.map((row) => ({
      Customer: row.customer,
      Method: row.method,
      Qty: row.qty,
      'Blank / Unit': row.blankCost,
      'Print / Unit': row.printMaterialCost,
      'Setup Hrs': row.setupHours,
      'Prod Hrs': row.productionHours,
      'QC Hrs': row.qcPackHours,
      Shipping: row.shipping,
      'Other Direct': row.otherDirect,
      'Rush?': row.rush ? 'Yes' : 'No',
      'Quoted Price': row.quotedPrice,
      'Total Cost': row.result?.totalJobCost || 0,
      'Recommended Price': row.result?.recommendedPrice || 0,
      'Actual Margin': row.result?.actualMarginPercent || 0,
    }))

    const quickQuoteSheet = [
      { Label: 'Quantity', Value: quickQuote.quantity },
      { Label: 'Blank Cost / Unit', Value: quickQuote.blankCostPerUnit },
      { Label: 'Decoration Cost / Unit', Value: quickQuote.decorationCostPerUnit },
      { Label: 'Setup + Artwork', Value: quickQuote.setupArtwork },
      { Label: 'Production + Overhead', Value: quickQuote.productionOverhead },
      { Label: 'Shipping / Other', Value: quickQuote.shippingOther },
      { Label: 'Target Margin %', Value: quickQuote.targetMarginPercent },
      { Label: 'Estimated Total Cost', Value: quickQuoteResults.totalCost },
      { Label: 'Cost per Unit', Value: quickQuoteResults.costPerUnit },
      { Label: 'Recommended Selling Price', Value: quickQuoteResults.recommendedSellingPrice },
      { Label: 'Recommended Price / Unit', Value: quickQuoteResults.recommendedPricePerUnit },
    ]

    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(assumptionsSheet), 'Assumptions')
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(machineRates), 'MachineRates')
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(jobs), 'Jobs')
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(quickQuoteSheet), 'QuickQuote')

    return workbook
  }

  function applyImportedAData(payload) {
    const nextCostState = {
      ...costState,
      jobCustomer: payload.costState?.jobCustomer ?? costState.jobCustomer,
      product: payload.costState?.product ?? costState.product,
      productionMethod: payload.costState?.productionMethod ?? costState.productionMethod,
      quantity: asNumber(payload.costState?.quantity ?? costState.quantity),
      sidesPrinted: asNumber(payload.costState?.sidesPrinted ?? costState.sidesPrinted),
      colorComplexity: asNumber(payload.costState?.colorComplexity ?? costState.colorComplexity),
      targetGrossMarginPercent: asNumber(payload.costState?.targetGrossMarginPercent ?? costState.targetGrossMarginPercent),
      testPricePerItem: asNumber(payload.costState?.testPricePerItem ?? costState.testPricePerItem),
      directRows: payload.directRows?.length ? payload.directRows : costState.directRows,
      laborRows: payload.laborRows?.length ? payload.laborRows : costState.laborRows,
    }

    setCostState(nextCostState)
    setActiveTab('cost')
  }

  function applyImportedBData(payload) {
    const nextAssumptions = {
      ...assumptions,
      ...payload.assumptions,
    }

    if (payload.machineRates) {
      nextAssumptions.machineRates = {
        ...assumptions.machineRates,
        ...payload.machineRates,
      }
    }

    setAssumptions(nextAssumptions)
    setJobRows(payload.jobRows?.length ? payload.jobRows : jobRows)
    setQuickQuote({ ...quickQuote, ...payload.quickQuote })
    setActiveTab('job')
  }

  function parseWorkbookImport(workbook, side) {
    const asRows = (sheet) => {
      if (!sheet) return []
      return XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false })
    }

    if (side === 'cost') {
      const settingsSheet = getSheetByName(workbook, ['Settings'])
      const directSheet = getSheetByName(workbook, ['DirectCosts'])
      const laborSheet = getSheetByName(workbook, ['LaborCosts'])

      if (!settingsSheet || !directSheet || !laborSheet) {
        return null
      }

      const settings = asRows(settingsSheet)
      const directRows = asRows(directSheet)
      const laborRows = asRows(laborSheet)

      const settingsMap = Object.fromEntries(
        settings.map((row) => [normalizeKey(row.key ?? row.Key ?? row.field), row.value ?? row.Value ?? row.val ?? '']),
      )

      const parsedDirectRows = directRows.map((row) => ({
        key: normalizeKey(row.name ?? row.Name ?? row.cost_component ?? row['Cost Component']) || 'custom',
        name: row.name ?? row.Name ?? row.cost_component ?? row['Cost Component'] ?? 'Cost item',
        unitCost: asNumber(row.unitCost ?? row['Unit Cost'] ?? row.unit_cost ?? 0),
        qty: asNumber(row.qty ?? row.Qty ?? row.quantity ?? 0),
        include: asBoolean(row.include ?? row.Include ?? row.included ?? true),
      }))

      const parsedLaborRows = laborRows.map((row) => ({
        key: normalizeKey(row.name ?? row.Name ?? row.cost_driver ?? row['Cost Driver']) || 'custom',
        name: row.name ?? row.Name ?? row.cost_driver ?? row['Cost Driver'] ?? 'Labor item',
        rateOrPercent: asNumber(row.rateOrPercent ?? row['Rate / %'] ?? row.rate ?? 0),
        baseHours: asNumber(row.baseHours ?? row['Hours / Base'] ?? row.base_hours ?? 0),
        include: asBoolean(row.include ?? row.Include ?? row.included ?? true),
      }))

      return {
        costState: {
          jobCustomer: settingsMap.jobcustomer || '',
          product: settingsMap.product || '',
          productionMethod: settingsMap.productionmethod || '',
          quantity: asNumber(settingsMap.quantity || 0),
          sidesPrinted: asNumber(settingsMap.sidesprinted || 0),
          colorComplexity: asNumber(settingsMap.colorcomplexity || 0),
          targetGrossMarginPercent: asNumber(settingsMap.targetgrossmarginpercent || 0),
          testPricePerItem: asNumber(settingsMap.testpriceperitem || 0),
        },
        directRows: parsedDirectRows,
        laborRows: parsedLaborRows,
      }
    }

    const assumptionsSheet = getSheetByName(workbook, ['Assumptions'])
    const machineRatesSheet = getSheetByName(workbook, ['MachineRates'])
    const jobsSheet = getSheetByName(workbook, ['Jobs'])
    const quickQuoteSheet = getSheetByName(workbook, ['QuickQuote'])

    if (!assumptionsSheet || !jobsSheet) {
      return null
    }

    const assumptionsRows = asRows(assumptionsSheet)
    const machineRows = asRows(machineRatesSheet)
    const jobRowsData = asRows(jobsSheet)
    const quickRows = asRows(quickQuoteSheet)

    const assumptionsMap = Object.fromEntries(
      assumptionsRows.map((row) => [normalizeKey(row.key ?? row.Key), row.value ?? row.Value ?? row.val ?? '']),
    )

    const machineRates = Object.fromEntries(
      machineRows.map((row) => [row.method ?? row.Method ?? row.name, asNumber(row.rate ?? row.Rate ?? row.value ?? 0)]),
    )

    const importedJobs = jobRowsData.map((row) => ({
      id: `job-import-${Math.random().toString(36).slice(2, 9)}`,
      customer: row.customer ?? row.Customer ?? '',
      method: row.method ?? row.Method ?? 'DTF / Heat Transfer',
      qty: asNumber(row.qty ?? row.Qty ?? 0),
      blankCost: asNumber(row.blankCost ?? row['Blank / Unit'] ?? row.blank ?? 0),
      printMaterialCost: asNumber(row.printMaterialCost ?? row['Print / Unit'] ?? row.print_material ?? 0),
      setupHours: asNumber(row.setupHours ?? row['Setup Hrs'] ?? row.setup ?? 0),
      productionHours: asNumber(row.productionHours ?? row['Prod Hrs'] ?? row.production ?? 0),
      qcPackHours: asNumber(row.qcPackHours ?? row['QC Hrs'] ?? row.qc ?? 0),
      shipping: asNumber(row.shipping ?? row.Shipping ?? 0),
      otherDirect: asNumber(row.otherDirect ?? row['Other Direct'] ?? row.other_direct ?? 0),
      rush: asBoolean(row.rush ?? row.Rush ?? row['Rush?'] ?? false),
      quotedPrice: asNumber(row.quotedPrice ?? row['Quoted Price'] ?? row.quoted ?? 0),
    }))

    const quickQuoteMap = Object.fromEntries(
      quickRows.map((row) => [normalizeKey(row.key ?? row.Key), row.value ?? row.Value ?? row.val ?? '']),
    )

    return {
      assumptions: {
        hourlyProductionLaborRate: asNumber(assumptionsMap.hourlyproductionlaborrate || assumptions.hourlyProductionLaborRate),
        laborBurdenPercent: asNumber(assumptionsMap.laborburdenpercent || assumptions.laborBurdenPercent),
        monthlyFixedOverhead: asNumber(assumptionsMap.monthlyfixedoverhead || assumptions.monthlyFixedOverhead),
        monthlyProductiveShopHours: asNumber(assumptionsMap.monthlyproductiveshophours || assumptions.monthlyProductiveShopHours),
        targetGrossMarginPercent: asNumber(assumptionsMap.targetgrossmarginpercent || assumptions.targetGrossMarginPercent),
        defaultWastePercent: asNumber(assumptionsMap.defaultwastepercent || assumptions.defaultWastePercent),
        rushFeePercent: asNumber(assumptionsMap.rushfeepercent || assumptions.rushFeePercent),
        salesFeePercent: asNumber(assumptionsMap.salesfeepercent || assumptions.salesFeePercent),
        setupArtworkRate: asNumber(assumptionsMap.setupartworkrate || assumptions.setupArtworkRate),
        qcRate: asNumber(assumptionsMap.qcrate || assumptions.qcRate),
        machineRates: { ...assumptions.machineRates, ...machineRates },
      },
      jobRows: importedJobs,
      quickQuote: {
        quantity: asNumber(quickQuoteMap.quantity || quickQuote.quantity),
        blankCostPerUnit: asNumber(quickQuoteMap.blankcostperunit || quickQuote.blankCostPerUnit),
        decorationCostPerUnit: asNumber(quickQuoteMap.decorationcostperunit || quickQuote.decorationCostPerUnit),
        setupArtwork: asNumber(quickQuoteMap.setupartwork || quickQuote.setupArtwork),
        productionOverhead: asNumber(quickQuoteMap.productionoverhead || quickQuote.productionOverhead),
        shippingOther: asNumber(quickQuoteMap.shippingother || quickQuote.shippingOther),
        targetMarginPercent: asNumber(quickQuoteMap.targetmarginpercent || quickQuote.targetMarginPercent),
      },
    }
  }

  async function handleImportFile(event) {
    const file = event.target.files?.[0]
    if (!file) return

    const targetSide = pendingImportSide ?? 'cost'

    try {
      if (file.name.toLowerCase().endsWith('.pdf')) {
        const text = await extractPdfText(file)
        const parsed = parsePdfImportText(text, targetSide)

        if (!parsed) {
          alert('PDF import did not find recognizable calculator fields. Please use a text-based PDF or export a spreadsheet template first.')
          event.target.value = ''
          setPendingImportSide(null)
          return
        }

        if (targetSide === 'cost') {
          applyImportedAData(parsed)
        } else {
          applyImportedBData(parsed)
        }
      } else {
        const buffer = await file.arrayBuffer()
        const workbook = XLSX.read(buffer, { type: 'array' })
        const parsed = parseWorkbookImport(workbook, targetSide)

        if (!parsed) {
          alert('This file does not match the expected calculator format. Please use the exported spreadsheet template or a similar layout.')
          event.target.value = ''
          setPendingImportSide(null)
          return
        }

        if (targetSide === 'cost') {
          applyImportedAData(parsed)
        } else {
          applyImportedBData(parsed)
        }
      }

      alert(`${targetSide === 'cost' ? 'A side' : 'B side'} import completed.`)
    } catch (error) {
      console.error('Import failed', error)
      alert('The file could not be imported. Please check that it is a valid Excel, CSV, or readable PDF file.')
    } finally {
      event.target.value = ''
      setPendingImportSide(null)
    }
  }

  function exportWorkbook(side) {
    const workbook = side === 'cost' ? buildAExportWorkbook() : buildBExportWorkbook()
    const filename = side === 'cost' ? 'print-shop-a-side-template.xlsx' : 'print-shop-b-side-template.xlsx'
    XLSX.writeFile(workbook, filename)
  }

  function buildATemplateWorkbook() {
    const workbook = XLSX.utils.book_new()

    const settings = [
      { Label: 'Job / Customer', Value: '' },
      { Label: 'Product / Garment', Value: '' },
      { Label: 'Production Method', Value: '' },
      { Label: 'Quantity', Value: 0 },
      { Label: 'Sides Printed', Value: 0 },
      { Label: 'Colors / Complexity', Value: 0 },
      { Label: 'Target Gross Margin %', Value: 0 },
      { Label: 'Your Price / Item (test)', Value: 0 },
    ]

    const directRows = directCostDefaults.map((row) => ({
      'Cost Component': row.name,
      'Unit Cost': 0,
      Qty: 0,
      Include: row.include ? 'Yes' : 'No',
      'Extended Cost': 0,
    }))

    const laborRows = laborDefaults.map((row) => ({
      'Cost Driver': row.name,
      'Rate / %': 0,
      'Hours / Base': 0,
      Include: row.include ? 'Yes' : 'No',
      'Calculated Cost': 0,
    }))

    const resultRows = [
      { Label: 'Total True Job Cost', Value: 0 },
      { Label: 'True Cost Per Item', Value: 0 },
      { Label: 'Break-even Price / Item', Value: 0 },
      { Label: 'Target Selling Price / Item', Value: 0 },
      { Label: 'Recommended Job Total', Value: 0 },
      { Label: 'Expected Gross Profit', Value: 0 },
      { Label: 'Expected Gross Margin', Value: 0 },
      { Label: 'Quoted Job Total', Value: 0 },
      { Label: 'Payment Fee', Value: 0 },
      { Label: 'Profit After Costs', Value: 0 },
      { Label: 'Profit / Item', Value: 0 },
      { Label: 'Test Gross Margin', Value: 0 },
      { Label: 'Difference vs Target', Value: 0 },
      { Label: 'Status', Value: '' },
    ]

    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(settings), 'Settings')
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(directRows), 'DirectCosts')
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(laborRows), 'LaborCosts')
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(resultRows), 'Results')

    return workbook
  }

  function buildBTemplateWorkbook() {
    const workbook = XLSX.utils.book_new()

    const assumptionsSheet = [
      { Label: 'Hourly Production Labor Rate', Value: assumptionsDefault.hourlyProductionLaborRate },
      { Label: 'Labor Burden %', Value: assumptionsDefault.laborBurdenPercent },
      { Label: 'Monthly Fixed Overhead', Value: assumptionsDefault.monthlyFixedOverhead },
      { Label: 'Monthly Productive Shop Hours', Value: assumptionsDefault.monthlyProductiveShopHours },
      { Label: 'Target Gross Margin %', Value: assumptionsDefault.targetGrossMarginPercent },
      { Label: 'Default Waste %', Value: assumptionsDefault.defaultWastePercent },
      { Label: 'Rush Fee %', Value: assumptionsDefault.rushFeePercent },
      { Label: 'Sales / Payment Processing %', Value: assumptionsDefault.salesFeePercent },
      { Label: 'Setup / Artwork Rate per Hour', Value: assumptionsDefault.setupArtworkRate },
      { Label: 'QC Rate per Hour', Value: assumptionsDefault.qcRate },
    ]

    const machineRates = Object.entries(assumptionsDefault.machineRates).map(([method, rate]) => ({
      Method: method,
      Rate: rate,
    }))

    const jobs = Array.from({ length: 6 }, (_, i) => ({
      Customer: '',
      Method: i === 0 ? 'DTF / Heat Transfer' : '',
      Qty: 0,
      'Blank / Unit': 0,
      'Print / Unit': 0,
      'Setup Hrs': 0,
      'Prod Hrs': 0,
      'QC Hrs': 0,
      Shipping: 0,
      'Other Direct': 0,
      'Rush?': 'No',
      'Quoted Price': 0,
      'Total Cost': 0,
      'Recommended Price': 0,
      'Actual Margin': 0,
    }))

    const quickQuoteSheet = [
      { Label: 'Quantity', Value: 0 },
      { Label: 'Blank Cost / Unit', Value: 0 },
      { Label: 'Decoration Cost / Unit', Value: 0 },
      { Label: 'Setup + Artwork', Value: 0 },
      { Label: 'Production + Overhead', Value: 0 },
      { Label: 'Shipping / Other', Value: 0 },
      { Label: 'Target Margin %', Value: 0 },
      { Label: 'Estimated Total Cost', Value: 0 },
      { Label: 'Cost per Unit', Value: 0 },
      { Label: 'Recommended Selling Price', Value: 0 },
      { Label: 'Recommended Price / Unit', Value: 0 },
    ]

    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(assumptionsSheet), 'Assumptions')
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(machineRates), 'MachineRates')
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(jobs), 'Jobs')
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(quickQuoteSheet), 'QuickQuote')

    return workbook
  }

  function downloadTemplate(side, fileType) {
    const workbook = side === 'cost' ? buildATemplateWorkbook() : buildBTemplateWorkbook()
    const sideName = side === 'cost' ? 'a-side' : 'b-side'
    const filename = `print-shop-${sideName}-import-template.${fileType}`

    XLSX.writeFile(workbook, filename, {
      bookType: fileType,
    })
  }

  function exportCsv(side) {
    const workbook = side === 'cost' ? buildAExportWorkbook() : buildBExportWorkbook()
    const firstSheetName = workbook.SheetNames[0]
    const csv = XLSX.utils.sheet_to_csv(workbook.Sheets[firstSheetName])
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' })
    const url = URL.createObjectURL(blob)
    const link = document.createElement('a')
    link.href = url
    link.download = side === 'cost' ? 'print-shop-a-side.csv' : 'print-shop-b-side.csv'
    link.click()
    URL.revokeObjectURL(url)
  }

  function addTableToPdf(pdf, headers, rows, startY, pageWidth = 190) {
    const colWidths = headers.map((header) => {
      const labels = rows.map((row) => String(row[header.key] ?? '')).concat(header.label)
      const max = Math.max(...labels.map((value) => value.length))
      return Math.max(18, Math.min(38, max * 2.8))
    })

    const totalWidth = colWidths.reduce((sum, width) => sum + width, 0)
    const startX = 14
    let currentY = startY
    const gap = 6

    const drawRow = (rowEntries) => {
      const rowY = currentY
      let x = startX
      rowEntries.forEach((entry, index) => {
        pdf.setFontSize(8)
        pdf.text(String(entry), x + 2, rowY + 6)
        x += colWidths[index]
      })
      currentY += gap
    }

    drawRow(headers.map((header) => header.label))
    pdf.line(startX, currentY - 2, startX + totalWidth, currentY - 2)

    rows.forEach((row) => {
      if (currentY > 260) {
        pdf.addPage()
        currentY = 20
      }
      drawRow(headers.map((header) => row[header.key]))
    })

    return currentY + 8
  }

  function exportPdf(side) {
    const pdf = new jsPDF()
    const title = side === 'cost' ? 'Print Shop Calculator - A Side' : 'Print Shop Calculator - B Side'
    let y = 16

    pdf.setFont('helvetica', 'bold')
    pdf.setFontSize(16)
    pdf.text(title, 14, y)
    y += 10

    if (side === 'cost') {
      const settings = [
        ['Job / Customer', costState.jobCustomer || 'N/A'],
        ['Product / Garment', costState.product || 'N/A'],
        ['Production Method', costState.productionMethod || 'N/A'],
        ['Quantity', costState.quantity || 0],
        ['Sides Printed', costState.sidesPrinted || 0],
        ['Colors / Complexity', costState.colorComplexity || 0],
        ['Target Gross Margin %', `${costState.targetGrossMarginPercent || 0}%`],
        ['Your Price / Item (test)', money(asNumber(costState.testPricePerItem))],
      ]

      pdf.setFont('helvetica', 'normal')
      pdf.setFontSize(10)
      settings.forEach(([label, value]) => {
        if (y > 270) {
          pdf.addPage()
          y = 20
        }
        pdf.text(`${label}: ${value}`, 14, y)
        y += 7
      })

      y += 4
      pdf.setFont('helvetica', 'bold')
      pdf.setFontSize(12)
      pdf.text('Direct Costs', 14, y)
      y += 8
      pdf.setFont('helvetica', 'normal')
      pdf.setFontSize(8)

      const directRows = costState.directRows.map((row, index) => ({
        name: row.name,
        unitCost: row.unitCost,
        qty: row.qty,
        include: row.include ? 'Yes' : 'No',
        extended: costResults.directRows[index]?.extended || 0,
      }))

      y = addTableToPdf(
        pdf,
        [
          { key: 'name', label: 'Cost Component' },
          { key: 'unitCost', label: 'Unit Cost' },
          { key: 'qty', label: 'Qty' },
          { key: 'include', label: 'Include' },
          { key: 'extended', label: 'Extended' },
        ],
        directRows,
        y,
      )

      pdf.setFont('helvetica', 'bold')
      pdf.setFontSize(12)
      pdf.text('Labor, Overhead & Pricing Drivers', 14, y)
      y += 8
      pdf.setFont('helvetica', 'normal')
      pdf.setFontSize(8)

      const laborRows = costState.laborRows.map((row, index) => ({
        name: row.name,
        rateOrPercent: row.rateOrPercent,
        baseHours: row.baseHours,
        include: row.include ? 'Yes' : 'No',
        calculated: costResults.laborRows[index]?.calculated || 0,
      }))

      y = addTableToPdf(
        pdf,
        [
          { key: 'name', label: 'Cost Driver' },
          { key: 'rateOrPercent', label: 'Rate / %' },
          { key: 'baseHours', label: 'Hours / Base' },
          { key: 'include', label: 'Include' },
          { key: 'calculated', label: 'Calculated' },
        ],
        laborRows,
        y,
      )

      pdf.setFont('helvetica', 'bold')
      pdf.setFontSize(12)
      pdf.text('Results', 14, y)
      y += 8
      pdf.setFont('helvetica', 'normal')
      pdf.setFontSize(9)

      const resultRows = [
        ['Total True Job Cost', money(costResults.totalTrueJobCost)],
        ['True Cost Per Item', money(costResults.trueCostPerItem)],
        ['Break-even Price / Item', money(costResults.breakEvenPricePerItem)],
        ['Target Selling Price / Item', money(costResults.targetSellingPricePerItem)],
        ['Recommended Job Total', money(costResults.recommendedJobTotal)],
        ['Expected Gross Profit', money(costResults.expectedGrossProfit)],
        ['Expected Gross Margin', pct(costResults.expectedGrossMargin * 100)],
        ['Quoted Job Total', money(costResults.quotedJobTotal)],
        ['Payment Fee', money(costResults.paymentFee)],
        ['Profit After Costs', money(costResults.profitAfterCosts)],
        ['Profit / Item', money(costResults.profitPerItem)],
        ['Test Gross Margin', pct(costResults.grossMargin * 100)],
        ['Difference vs Target', money(costResults.differenceVsTarget)],
        ['Status', costResults.status],
      ]

      resultRows.forEach(([label, value]) => {
        if (y > 270) {
          pdf.addPage()
          y = 20
        }
        pdf.text(`${label}: ${value}`, 14, y)
        y += 7
      })
    } else {
      const assumptionsRows = [
        ['Hourly Production Labor Rate', money(asNumber(assumptions.hourlyProductionLaborRate))],
        ['Labor Burden %', `${assumptions.laborBurdenPercent}%`],
        ['Monthly Fixed Overhead', money(asNumber(assumptions.monthlyFixedOverhead))],
        ['Monthly Productive Shop Hours', assumptions.monthlyProductiveShopHours],
        ['Target Gross Margin %', `${assumptions.targetGrossMarginPercent}%`],
        ['Default Waste %', `${assumptions.defaultWastePercent}%`],
        ['Rush Fee %', `${assumptions.rushFeePercent}%`],
        ['Sales / Payment Processing %', `${assumptions.salesFeePercent}%`],
        ['Setup / Artwork Rate per Hour', money(asNumber(assumptions.setupArtworkRate))],
        ['QC Rate per Hour', money(asNumber(assumptions.qcRate))],
      ]

      pdf.setFont('helvetica', 'normal')
      pdf.setFontSize(10)
      assumptionsRows.forEach(([label, value]) => {
        if (y > 270) {
          pdf.addPage()
          y = 20
        }
        pdf.text(`${label}: ${value}`, 14, y)
        y += 7
      })

      y += 4
      pdf.setFont('helvetica', 'bold')
      pdf.setFontSize(12)
      pdf.text('Machine Rates by Method', 14, y)
      y += 8
      pdf.setFont('helvetica', 'normal')
      pdf.setFontSize(8)
      y = addTableToPdf(
        pdf,
        [
          { key: 'method', label: 'Method' },
          { key: 'rate', label: 'Rate' },
        ],
        Object.entries(assumptions.machineRates).map(([method, rate]) => ({ method, rate })),
        y,
      )

      pdf.setFont('helvetica', 'bold')
      pdf.setFontSize(12)
      pdf.text('Job Table', 14, y)
      y += 8
      pdf.setFont('helvetica', 'normal')
      pdf.setFontSize(8)

      y = addTableToPdf(
        pdf,
        [
          { key: 'customer', label: 'Customer' },
          { key: 'method', label: 'Method' },
          { key: 'qty', label: 'Qty' },
          { key: 'blankCost', label: 'Blank' },
          { key: 'printMaterialCost', label: 'Print' },
        ],
        jobRows.map((row) => ({
          customer: row.customer,
          method: row.method,
          qty: row.qty,
          blankCost: row.blankCost,
          printMaterialCost: row.printMaterialCost,
        })),
        y,
      )

      pdf.setFont('helvetica', 'bold')
      pdf.setFontSize(12)
      pdf.text('Quick Quote Calculator', 14, y)
      y += 8
      pdf.setFont('helvetica', 'normal')
      pdf.setFontSize(9)

      const quickRows = [
        ['Quantity', quickQuote.quantity],
        ['Blank Cost / Unit', money(asNumber(quickQuote.blankCostPerUnit))],
        ['Decoration Cost / Unit', money(asNumber(quickQuote.decorationCostPerUnit))],
        ['Setup + Artwork', money(asNumber(quickQuote.setupArtwork))],
        ['Production + Overhead', money(asNumber(quickQuote.productionOverhead))],
        ['Shipping / Other', money(asNumber(quickQuote.shippingOther))],
        ['Target Margin %', `${quickQuote.targetMarginPercent}%`],
        ['Estimated Total Cost', money(quickQuoteResults.totalCost)],
        ['Cost per Unit', money(quickQuoteResults.costPerUnit)],
        ['Recommended Selling Price', money(quickQuoteResults.recommendedSellingPrice)],
        ['Recommended Price / Unit', money(quickQuoteResults.recommendedPricePerUnit)],
      ]

      quickRows.forEach(([label, value]) => {
        if (y > 270) {
          pdf.addPage()
          y = 20
        }
        pdf.text(`${label}: ${value}`, 14, y)
        y += 7
      })
    }

    pdf.save(side === 'cost' ? 'print-shop-a-side.pdf' : 'print-shop-b-side.pdf')
  }

  return (
    <main className="app-shell">
      <div className="ambient ambient-left" aria-hidden="true"></div>
      <div className="ambient ambient-right" aria-hidden="true"></div>

      <header className="hero">
        <p className="eyebrow">Powered by 1st Step Branding</p>
        <h1>Print Shop Cost Calculators</h1>
        <p className="lede">
          React implementation of your two workbook tabs with equivalent formulas and no backend.
          Everything runs in the browser.
        </p>
      </header>

      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,.xls,.csv,.pdf"
        style={{ display: 'none' }}
        onChange={handleImportFile}
      />

      {importModalOpen && (
        <div className="modal-backdrop" onClick={() => setImportModalOpen(false)}>
          <div className="modal-card" onClick={(event) => event.stopPropagation()}>
            <h3>Which calculator are you importing into?</h3>
            <div className="modal-actions">
              <button type="button" className="action-button primary" onClick={() => triggerImport('cost')}>
                A side
              </button>
              <button type="button" className="action-button" onClick={() => triggerImport('job')}>
                B side
              </button>
            </div>
            <button type="button" className="text-button" onClick={() => setImportModalOpen(false)}>
              Cancel
            </button>
          </div>
        </div>
      )}

      {exportModalOpen && (
        <div className="modal-backdrop" onClick={() => setExportModalOpen(false)}>
          <div className="modal-card" onClick={(event) => event.stopPropagation()}>
            <h3>Choose export file type</h3>
            <div className="modal-actions column">
              <button
                type="button"
                className="action-button primary"
                onClick={() => {
                  exportWorkbook(activeTab === 'job' ? 'job' : 'cost')
                  setExportModalOpen(false)
                }}
              >
                Excel
              </button>
              <button
                type="button"
                className="action-button"
                onClick={() => {
                  exportCsv(activeTab === 'job' ? 'job' : 'cost')
                  setExportModalOpen(false)
                }}
              >
                CSV
              </button>
              <button
                type="button"
                className="action-button"
                onClick={() => {
                  exportPdf(activeTab === 'job' ? 'job' : 'cost')
                  setExportModalOpen(false)
                }}
              >
                PDF
              </button>
              <button
                type="button"
                className="action-button"
                onClick={() => {
                  setTemplateSide(activeTab === 'job' ? 'job' : 'cost')
                  setExportModalOpen(false)
                  setTemplateModalOpen(true)
                }}
              >
                Template
              </button>
            </div>
            <button type="button" className="text-button" onClick={() => setExportModalOpen(false)}>
              Cancel
            </button>
          </div>
        </div>
      )}

      {templateModalOpen && (
        <div className="modal-backdrop" onClick={() => setTemplateModalOpen(false)}>
          <div className="modal-card" onClick={(event) => event.stopPropagation()}>
            <h3>Download import template</h3>
            <p>Select the calculator side and file type for your editable template.</p>
            <div className="modal-actions">
              <button
                type="button"
                className={templateSide === 'cost' ? 'action-button primary' : 'action-button'}
                onClick={() => setTemplateSide('cost')}
              >
                A side template
              </button>
              <button
                type="button"
                className={templateSide === 'job' ? 'action-button primary' : 'action-button'}
                onClick={() => setTemplateSide('job')}
              >
                B side template
              </button>
            </div>
            <div className="modal-actions">
              <button
                type="button"
                className="action-button primary"
                onClick={() => {
                  downloadTemplate(templateSide, 'xlsx')
                  setTemplateModalOpen(false)
                }}
              >
                Excel (.xlsx)
              </button>
              <button
                type="button"
                className="action-button"
                onClick={() => {
                  downloadTemplate(templateSide, 'xls')
                  setTemplateModalOpen(false)
                }}
              >
                Excel 97-2003 (.xls)
              </button>
            </div>
            <button type="button" className="text-button" onClick={() => setTemplateModalOpen(false)}>
              Cancel
            </button>
          </div>
        </div>
      )}

      <section className="tabs" aria-label="Calculator tabs">
        <div className="tabs-left">
          <button
            type="button"
            className={activeTab === 'cost' ? 'tab active' : 'tab'}
            onClick={() => setActiveTab('cost')}
          >
            Job Cost Calculator A side
          </button>
          <button
            type="button"
            className={activeTab === 'job' ? 'tab active' : 'tab'}
            onClick={() => setActiveTab('job')}
          >
            Job Cost Calculator B side
          </button>
          <button
            type="button"
            className={activeTab === 'guide' ? 'tab active' : 'tab'}
            onClick={() => setActiveTab('guide')}
          >
            How to Use
          </button>
        </div>
        <div className="tabs-right">
          <button type="button" className="action-button primary" onClick={() => setImportModalOpen(true)}>
            Import
          </button>
          <button type="button" className="action-button" onClick={() => setExportModalOpen(true)}>
            Export
          </button>
        </div>
      </section>

      {activeTab === 'guide' && (
        <section className="panel">
          <div className="panel-header">
            <h2>How to Use the Calculators</h2>
          </div>

          <div className="guide-tabs" aria-label="Usage guide tabs">
            <button
              type="button"
              className={guideFocus === 'cost' ? 'guide-pill active' : 'guide-pill'}
              onClick={() => setGuideFocus('cost')}
            >
              A Side How To
            </button>
            <button
              type="button"
              className={guideFocus === 'job' ? 'guide-pill active' : 'guide-pill'}
              onClick={() => setGuideFocus('job')}
            >
              B Side How To
            </button>
          </div>

          {guideFocus === 'cost' ? (
            <div className="help-content">
              <h3>What the A side calculator is used for</h3>
              <p>
                The A side calculator is the main pricing tool for a new job. It helps you estimate
                the real cost of producing a run, understand your break-even point, and test what
                selling price will reach your target margin before you quote a customer.
              </p>

              <h3>1. Job and production details</h3>
              <ul>
                <li>
                  <strong>Job / Customer:</strong> Enter the customer name or job label so you can
                  keep track of what you are pricing.
                </li>
                <li>
                  <strong>Product / Garment:</strong> Name the garment or item being produced, such as
                  a T-shirt, hoodie, tote, or different apparel style.
                </li>
                <li>
                  <strong>Production Method:</strong> Choose the decorating process you are using,
                  such as DTF, screen printing, HTV, embroidery, or sublimation.
                </li>
                <li>
                  <strong>Quantity:</strong> Enter the total number of units in the order. This drives
                  the cost per piece and overall job total.
                </li>
                <li>
                  <strong>Sides Printed:</strong> Enter how many sides are being printed, such as 1 for
                  front only or 2 for front and back.
                </li>
                <li>
                  <strong>Colors / Complexity:</strong> This is a rough complexity factor for the job.
                  More colors, harder artwork, or more detailed production generally raises the cost.
                </li>
              </ul>

              <h3>2. Direct costs</h3>
              <p>
                This section covers all the costs that are directly tied to the job. Each row shows a
                cost component, the cost per unit, the quantity, whether it is included, and the
                extended cost.
              </p>
              <ul>
                <li>
                  <strong>Blank garment / product:</strong> The base product cost for each unit.
                </li>
                <li>
                  <strong>Transfer / ink / vinyl / thread:</strong> Material costs tied to the print
                  method or application process.
                </li>
                <li>
                  <strong>Packaging:</strong> Bags, labels, boxes, or protective packaging per unit.
                </li>
                <li>
                  <strong>Outside services:</strong> Any outsourced item like fulfillment or outside
                  labor.
                </li>
                <li>
                  <strong>Shipping / freight:</strong> Logistics cost for shipping materials or delivery.
                </li>
                <li>
                  <strong>Setup / artwork:</strong> One-time design or setup costs for a run.
                </li>
                <li>
                  <strong>Other direct cost:</strong> Any additional direct expense not listed above.
                </li>
              </ul>

              <h3>3. Labor, overhead, and pricing drivers</h3>
              <ul>
                <li>
                  <strong>Rate / %:</strong> This is the labor rate or percentage used to calculate the
                  cost for that category.
                </li>
                <li>
                  <strong>Hours / Base:</strong> Enter the estimated hours for labor rows, or a base
                  number when the formula depends on a percentage of direct cost.
                </li>
                <li>
                  <strong>Include:</strong> Turn this off if a cost driver should not be included in the
                  estimate.
                </li>
                <li>
                  <strong>Calculated Cost:</strong> The system automatically creates this value using the
                  formula tied to the row.
                </li>
                <li>
                  <strong>Target Gross Margin %:</strong> This is your desired profit target. A higher
                  percentage means you want more profit left after costs.
                </li>
                <li>
                  <strong>Your Price / Item (test):</strong> This is a temporary pricing check, not the
                  final customer quote. Enter a test selling price to see whether your job meets the
                  target margin before you quote a customer. If the price is high enough, the status
                  will show “Meets target.” If it is too low, it will show “Below target.”
                </li>
              </ul>

              <h3>How to read the A side results</h3>
              <p>
                Review the total true job cost, true cost per item, break-even price, target selling
                price, and expected gross margin. If the projected margin is below your goal, adjust
                the selling price, quantity, or cost inputs until the estimate aligns with your desired
                profitability.
              </p>
            </div>
          ) : (
            <div className="help-content">
              <h3>What the B side calculator is used for</h3>
              <p>
                The B side calculator is best used for job-by-job production planning and quote
                tracking. It lets you enter the real assumptions for each order, compare quoted price
                to actual cost, and see whether each job is meeting your target margin.
              </p>

              <h3>1. Rates and assumptions</h3>
              <ul>
                <li>
                  <strong>Hourly Production Labor Rate:</strong> Your labor cost per hour for production
                  work.
                </li>
                <li>
                  <strong>Labor Burden %:</strong> Additional payroll costs such as taxes, benefits,
                  insurance, and payroll extras as a percentage of direct labor.
                </li>
                <li>
                  <strong>Monthly Fixed Overhead:</strong> Monthly shop expenses like rent, utilities,
                  software, and admin costs divided into the cost model.
                </li>
                <li>
                  <strong>Monthly Productive Shop Hours:</strong> The number of billable or productive
                  hours your shop can realistically deliver each month.
                </li>
                <li>
                  <strong>Target Gross Margin %:</strong> The profit goal for jobs and quotes.
                </li>
                <li>
                  <strong>Default Waste %:</strong> A percentage used for spoilage or excess material.
                </li>
                <li>
                  <strong>Rush Fee %:</strong> Extra pricing added when a customer needs expedited work.
                </li>
                <li>
                  <strong>Sales / Payment Processing %:</strong> The fee tied to processing the sale or
                  sales commission factor used in pricing.
                </li>
                <li>
                  <strong>Setup / Artwork Rate per Hour:</strong> Cost per hour for design, prepress,
                  setup, or artwork adjustments.
                </li>
                <li>
                  <strong>QC Rate per Hour:</strong> Cost per hour for quality checks, finishing, and
                  final inspection.
                </li>
              </ul>

              <h3>2. Machine rates by production method</h3>
              <p>
                Each production method has its own machine rate. This represents the hourly machine or
                equipment cost tied to the process. For example, embroidery and DTG usually carry
                different equipment costs than HTV or screen printing.
              </p>

              <h3>3. Job table</h3>
              <ul>
                <li>
                  <strong>Customer:</strong> Name of the customer or job.
                </li>
                <li>
                  <strong>Method:</strong> Which production method is used for that job.
                </li>
                <li>
                  <strong>Qty:</strong> Number of units in the job.
                </li>
                <li>
                  <strong>Blank / Unit:</strong> Cost of the blank or garment for each unit.
                </li>
                <li>
                  <strong>Print / Unit:</strong> Material cost for decoration per unit.
                </li>
                <li>
                  <strong>Setup Hrs:</strong> Time required to prepare the job before production begins.
                </li>
                <li>
                  <strong>Prod Hrs:</strong> Time spent actually producing the job.
                </li>
                <li>
                  <strong>QC Hrs:</strong> Time spent on quality control, inspection, and packing.
                </li>
                <li>
                  <strong>Shipping:</strong> Freight or shipping cost for the order.
                </li>
                <li>
                  <strong>Other Direct:</strong> Any extra direct cost for that specific order.
                </li>
                <li>
                  <strong>Rush?:</strong> Check this box if the customer needs expedited handling.
                </li>
                <li>
                  <strong>Quoted Price:</strong> The price you quoted the customer for the job.
                </li>
              </ul>

              <h3>4. Quick quote calculator</h3>
              <p>
                This section gives you a simplified estimate for a one-off quote without filling in the
                whole job table. It is useful when you want to quickly check the minimum price needed to
                maintain a target margin.
              </p>
              <ul>
                <li>
                  <strong>Quantity:</strong> Units in the quick quote.
                </li>
                <li>
                  <strong>Blank Cost / Unit:</strong> Base material cost per item.
                </li>
                <li>
                  <strong>Decoration Cost / Unit:</strong> Cost to decorate each item.
                </li>
                <li>
                  <strong>Setup + Artwork:</strong> One-time design or setup cost.
                </li>
                <li>
                  <strong>Production + Overhead:</strong> Added cost for labor, overhead, and run costs.
                </li>
                <li>
                  <strong>Shipping / Other:</strong> Any additional freight or miscellaneous direct cost.
                </li>
                <li>
                  <strong>Target Margin %:</strong> Desired profit to maintain on the quote.
                </li>
              </ul>

              <h3>How to read the B side results</h3>
              <p>
                The table calculates total cost, recommended price, and actual margin for each order.
                Use this page to compare your quoted price to real cost, identify underpriced jobs, and
                see whether the job is hitting the profit target before you accept or complete the order.
              </p>
            </div>
          )}
        </section>
      )}

      {activeTab === 'cost' && (
        <section className="panel">
          <h2>Cost Calculator</h2>

          <div className="grid two-col">
            <label>
              Job / Customer
              <input
                value={costState.jobCustomer}
                onChange={(e) => updateCostField('jobCustomer', e.target.value)}
              />
            </label>
            <label>
              Product / Garment
              <input
                value={costState.product}
                onChange={(e) => updateCostField('product', e.target.value)}
              />
            </label>
            <label>
              Production Method
              <input
                value={costState.productionMethod}
                onChange={(e) => updateCostField('productionMethod', e.target.value)}
              />
            </label>
            <label>
              Quantity
              <input
                type="number"
                min="0"
                value={costState.quantity}
                onChange={(e) => updateCostField('quantity', normalizeNumericInput(e.target.value))}
              />
            </label>
            <label>
              Sides Printed
              <input
                type="number"
                min="0"
                value={costState.sidesPrinted}
                onChange={(e) => updateCostField('sidesPrinted', normalizeNumericInput(e.target.value))}
              />
            </label>
            <label>
              Colors / Complexity
              <input
                type="number"
                min="0"
                value={costState.colorComplexity}
                onChange={(e) => updateCostField('colorComplexity', normalizeNumericInput(e.target.value))}
              />
            </label>
          </div>

          <h3>Direct Costs</h3>
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Cost Component</th>
                  <th>Unit Cost</th>
                  <th>Qty</th>
                  <th>Include</th>
                  <th>Extended Cost</th>
                </tr>
              </thead>
              <tbody>
                {costState.directRows.map((row, index) => (
                  <tr key={row.key}>
                    <td>{row.name}</td>
                    <td>
                      <input
                        type="number"
                        step="0.01"
                        value={row.unitCost}
                        onChange={(e) =>
                          updateDirectRow(index, 'unitCost', normalizeNumericInput(e.target.value))
                        }
                      />
                    </td>
                    <td>
                      <input
                        type="number"
                        step="0.01"
                        value={row.qty}
                        onChange={(e) => updateDirectRow(index, 'qty', normalizeNumericInput(e.target.value))}
                      />
                    </td>
                    <td>
                      <input
                        type="checkbox"
                        checked={row.include}
                        onChange={(e) => updateDirectRow(index, 'include', e.target.checked)}
                      />
                    </td>
                    <td>{money(costResults.directRows[index]?.extended || 0)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <h3>Labor, Overhead, and Pricing Drivers</h3>
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Cost Driver</th>
                  <th>Rate / %</th>
                  <th>Hours / Base</th>
                  <th>Include</th>
                  <th>Calculated Cost</th>
                </tr>
              </thead>
              <tbody>
                {costState.laborRows.map((row, index) => (
                  <tr key={row.key}>
                    <td>{row.name}</td>
                    <td>
                      <input
                        type="number"
                        step="0.01"
                        value={row.rateOrPercent}
                        onChange={(e) =>
                          updateLaborRow(index, 'rateOrPercent', normalizeNumericInput(e.target.value))
                        }
                      />
                    </td>
                    <td>
                      {index <= 3 ? (
                        <input
                          type="number"
                          step="0.01"
                          value={row.baseHours}
                          onChange={(e) =>
                            updateLaborRow(index, 'baseHours', normalizeNumericInput(e.target.value))
                          }
                        />
                      ) : (
                        <span className="muted">Formula-driven</span>
                      )}
                    </td>
                    <td>
                      <input
                        type="checkbox"
                        checked={row.include}
                        onChange={(e) => updateLaborRow(index, 'include', e.target.checked)}
                      />
                    </td>
                    <td>{money(costResults.laborRows[index]?.calculated || 0)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div className="grid two-col">
            <label>
              Target Gross Margin %
              <input
                type="number"
                step="0.1"
                value={costState.targetGrossMarginPercent}
                onChange={(e) =>
                  updateCostField('targetGrossMarginPercent', normalizeNumericInput(e.target.value))
                }
              />
            </label>
            <label>
              Your Price / Item (test)
              <input
                type="number"
                step="0.01"
                value={costState.testPricePerItem}
                onChange={(e) => updateCostField('testPricePerItem', normalizeNumericInput(e.target.value))}
              />
            </label>
          </div>

          <div className="results-grid">
            <ResultCard label="Total True Job Cost" value={money(costResults.totalTrueJobCost)} />
            <ResultCard label="True Cost Per Item" value={money(costResults.trueCostPerItem)} />
            <ResultCard
              label="Break-even Price / Item"
              value={money(costResults.breakEvenPricePerItem)}
            />
            <ResultCard
              label="Target Selling Price / Item"
              value={money(costResults.targetSellingPricePerItem)}
            />
            <ResultCard label="Recommended Job Total" value={money(costResults.recommendedJobTotal)} />
            <ResultCard label="Expected Gross Profit" value={money(costResults.expectedGrossProfit)} />
            <ResultCard
              label="Expected Gross Margin"
              value={pct(costResults.expectedGrossMargin * 100)}
            />
            <ResultCard label="Quoted Job Total" value={money(costResults.quotedJobTotal)} />
            <ResultCard label="Payment Fee" value={money(costResults.paymentFee)} />
            <ResultCard label="Profit After Costs" value={money(costResults.profitAfterCosts)} />
            <ResultCard label="Profit / Item" value={money(costResults.profitPerItem)} />
            <ResultCard label="Test Gross Margin" value={pct(costResults.grossMargin * 100)} />
            <ResultCard label="Difference vs Target" value={money(costResults.differenceVsTarget)} />
            <ResultCard label="Status" value={costResults.status} />
          </div>
        </section>
      )}

      {activeTab === 'job' && (
        <section className="panel">
          <h2>Job Cost Calculator</h2>

          <h3>Rates and Assumptions</h3>
          <div className="grid three-col">
            <label>
              Hourly Production Labor Rate
              <input
                type="number"
                step="0.01"
                value={assumptions.hourlyProductionLaborRate}
                onChange={(e) =>
                  updateAssumption('hourlyProductionLaborRate', normalizeNumericInput(e.target.value))
                }
              />
            </label>
            <label>
              Labor Burden %
              <input
                type="number"
                step="0.1"
                value={assumptions.laborBurdenPercent}
                onChange={(e) => updateAssumption('laborBurdenPercent', normalizeNumericInput(e.target.value))}
              />
            </label>
            <label>
              Monthly Fixed Overhead
              <input
                type="number"
                step="0.01"
                value={assumptions.monthlyFixedOverhead}
                onChange={(e) => updateAssumption('monthlyFixedOverhead', normalizeNumericInput(e.target.value))}
              />
            </label>
            <label>
              Monthly Productive Shop Hours
              <input
                type="number"
                step="0.01"
                value={assumptions.monthlyProductiveShopHours}
                onChange={(e) =>
                  updateAssumption('monthlyProductiveShopHours', normalizeNumericInput(e.target.value))
                }
              />
            </label>
            <label>
              Target Gross Margin %
              <input
                type="number"
                step="0.1"
                value={assumptions.targetGrossMarginPercent}
                onChange={(e) =>
                  updateAssumption('targetGrossMarginPercent', normalizeNumericInput(e.target.value))
                }
              />
            </label>
            <label>
              Default Waste %
              <input
                type="number"
                step="0.1"
                value={assumptions.defaultWastePercent}
                onChange={(e) => updateAssumption('defaultWastePercent', normalizeNumericInput(e.target.value))}
              />
            </label>
            <label>
              Rush Fee %
              <input
                type="number"
                step="0.1"
                value={assumptions.rushFeePercent}
                onChange={(e) => updateAssumption('rushFeePercent', normalizeNumericInput(e.target.value))}
              />
            </label>
            <label>
              Sales/Payment Processing %
              <input
                type="number"
                step="0.1"
                value={assumptions.salesFeePercent}
                onChange={(e) => updateAssumption('salesFeePercent', normalizeNumericInput(e.target.value))}
              />
            </label>
            <label>
              Setup/Artwork Rate per Hour
              <input
                type="number"
                step="0.01"
                value={assumptions.setupArtworkRate}
                onChange={(e) => updateAssumption('setupArtworkRate', normalizeNumericInput(e.target.value))}
              />
            </label>
            <label>
              QC Rate per Hour
              <input
                type="number"
                step="0.01"
                value={assumptions.qcRate}
                onChange={(e) => updateAssumption('qcRate', normalizeNumericInput(e.target.value))}
              />
            </label>
          </div>

          <h3>Machine Rates by Method</h3>
          <div className="grid three-col">
            {productionMethods.map((method) => (
              <label key={method}>
                {method}
                <input
                  type="number"
                  step="0.01"
                  value={assumptions.machineRates[method]}
                  onChange={(e) => updateMachineRate(method, normalizeNumericInput(e.target.value))}
                />
              </label>
            ))}
          </div>

          <h3>Job Table</h3>
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Customer</th>
                  <th>Method</th>
                  <th>Qty</th>
                  <th>Blank / Unit</th>
                  <th>Print / Unit</th>
                  <th>Setup Hrs</th>
                  <th>Prod Hrs</th>
                  <th>QC Hrs</th>
                  <th>Shipping</th>
                  <th>Other Direct</th>
                  <th>Rush?</th>
                  <th>Quoted Price</th>
                  <th>Total Cost</th>
                  <th>Recommended Price</th>
                  <th>Actual Margin</th>
                </tr>
              </thead>
              <tbody>
                {calculatedRows.map((row) => (
                  <tr key={row.id}>
                    <td>
                      <input
                        value={row.customer}
                        onChange={(e) => updateJobRow(row.id, 'customer', e.target.value)}
                      />
                    </td>
                    <td>
                      <select
                        value={row.method}
                        onChange={(e) => updateJobRow(row.id, 'method', e.target.value)}
                      >
                        {productionMethods.map((method) => (
                          <option key={method} value={method}>
                            {method}
                          </option>
                        ))}
                      </select>
                    </td>
                    <td>
                      <input
                        type="number"
                        value={row.qty}
                        onChange={(e) => updateJobRow(row.id, 'qty', normalizeNumericInput(e.target.value))}
                      />
                    </td>
                    <td>
                      <input
                        type="number"
                        step="0.01"
                        value={row.blankCost}
                        onChange={(e) =>
                          updateJobRow(row.id, 'blankCost', normalizeNumericInput(e.target.value))
                        }
                      />
                    </td>
                    <td>
                      <input
                        type="number"
                        step="0.01"
                        value={row.printMaterialCost}
                        onChange={(e) =>
                          updateJobRow(row.id, 'printMaterialCost', normalizeNumericInput(e.target.value))
                        }
                      />
                    </td>
                    <td>
                      <input
                        type="number"
                        step="0.01"
                        value={row.setupHours}
                        onChange={(e) =>
                          updateJobRow(row.id, 'setupHours', normalizeNumericInput(e.target.value))
                        }
                      />
                    </td>
                    <td>
                      <input
                        type="number"
                        step="0.01"
                        value={row.productionHours}
                        onChange={(e) =>
                          updateJobRow(row.id, 'productionHours', normalizeNumericInput(e.target.value))
                        }
                      />
                    </td>
                    <td>
                      <input
                        type="number"
                        step="0.01"
                        value={row.qcPackHours}
                        onChange={(e) =>
                          updateJobRow(row.id, 'qcPackHours', normalizeNumericInput(e.target.value))
                        }
                      />
                    </td>
                    <td>
                      <input
                        type="number"
                        step="0.01"
                        value={row.shipping}
                        onChange={(e) =>
                          updateJobRow(row.id, 'shipping', normalizeNumericInput(e.target.value))
                        }
                      />
                    </td>
                    <td>
                      <input
                        type="number"
                        step="0.01"
                        value={row.otherDirect}
                        onChange={(e) =>
                          updateJobRow(row.id, 'otherDirect', normalizeNumericInput(e.target.value))
                        }
                      />
                    </td>
                    <td>
                      <input
                        type="checkbox"
                        checked={row.rush}
                        onChange={(e) => updateJobRow(row.id, 'rush', e.target.checked)}
                      />
                    </td>
                    <td>
                      <input
                        type="number"
                        step="0.01"
                        value={row.quotedPrice}
                        onChange={(e) =>
                          updateJobRow(row.id, 'quotedPrice', normalizeNumericInput(e.target.value))
                        }
                      />
                    </td>
                    <td>{money(row.result.totalJobCost)}</td>
                    <td>{money(row.result.recommendedPrice)}</td>
                    <td>{pct(row.result.actualMarginPercent * 100)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <button type="button" className="add-row" onClick={addJobRow}>
            Add Job Row
          </button>

          <h3>Quick Quote Calculator</h3>
          <div className="grid two-col">
            <label>
              Quantity
              <input
                type="number"
                value={quickQuote.quantity}
                onChange={(e) => updateQuickQuote('quantity', normalizeNumericInput(e.target.value))}
              />
            </label>
            <label>
              Blank Cost / Unit
              <input
                type="number"
                step="0.01"
                value={quickQuote.blankCostPerUnit}
                onChange={(e) => updateQuickQuote('blankCostPerUnit', normalizeNumericInput(e.target.value))}
              />
            </label>
            <label>
              Decoration Cost / Unit
              <input
                type="number"
                step="0.01"
                value={quickQuote.decorationCostPerUnit}
                onChange={(e) =>
                  updateQuickQuote('decorationCostPerUnit', normalizeNumericInput(e.target.value))
                }
              />
            </label>
            <label>
              Setup + Artwork
              <input
                type="number"
                step="0.01"
                value={quickQuote.setupArtwork}
                onChange={(e) => updateQuickQuote('setupArtwork', normalizeNumericInput(e.target.value))}
              />
            </label>
            <label>
              Production + Overhead
              <input
                type="number"
                step="0.01"
                value={quickQuote.productionOverhead}
                onChange={(e) => updateQuickQuote('productionOverhead', normalizeNumericInput(e.target.value))}
              />
            </label>
            <label>
              Shipping / Other
              <input
                type="number"
                step="0.01"
                value={quickQuote.shippingOther}
                onChange={(e) => updateQuickQuote('shippingOther', normalizeNumericInput(e.target.value))}
              />
            </label>
            <label>
              Target Margin %
              <input
                type="number"
                step="0.1"
                value={quickQuote.targetMarginPercent}
                onChange={(e) => updateQuickQuote('targetMarginPercent', normalizeNumericInput(e.target.value))}
              />
            </label>
          </div>

          <div className="results-grid">
            <ResultCard label="Estimated Total Cost" value={money(quickQuoteResults.totalCost)} />
            <ResultCard label="Cost per Unit" value={money(quickQuoteResults.costPerUnit)} />
            <ResultCard
              label="Recommended Selling Price"
              value={money(quickQuoteResults.recommendedSellingPrice)}
            />
            <ResultCard
              label="Recommended Price / Unit"
              value={money(quickQuoteResults.recommendedPricePerUnit)}
            />
          </div>
        </section>
      )}
      <p className="version-badge" aria-label="Application version">
        Version 1.0
      </p>
    </main>
  )
}

function ResultCard({ label, value }) {
  return (
    <article className="result-card">
      <p>{label}</p>
      <strong>{value}</strong>
    </article>
  )
}

export default App

// Test de normalización
const normalizeTargetCode = (targetCode) => {
  if (!targetCode || typeof targetCode !== 'string') return targetCode
  let cleaned = targetCode.trim()
  
  const medidaMatch = cleaned.match(/Target\s+N\/A\s*\(\s*Medida\s+(\d+)\s*\)/i)
  if (medidaMatch) return `Target N° ${medidaMatch[1]}`
  
  const naMatch = cleaned.match(/N\/A\s*\(\s*Medida\s+(\d+)\s*\)/i)
  if (naMatch) return `Target N° ${naMatch[1]}`
  
  const pillarMatch = cleaned.match(/Target\s+N\/A\s*\(\s*Pilar\s+(\d+)\s*\)/i)
  if (pillarMatch) return `Target N° ${pillarMatch[1]}`
  
  return cleaned
}

const tests = [
  { input: 'Target N/A (Medida 1)', expected: 'Target N° 1' },
  { input: 'Target N/A (Medida 2)', expected: 'Target N° 2' },
  { input: 'Target N/A (Medida 3)', expected: 'Target N° 3' },
  { input: 'N/A (Medida 1)', expected: 'Target N° 1' },
  { input: 'Target N° 1', expected: 'Target N° 1' },
]

console.log('\n' + '='.repeat(50))
console.log('PRUEBAS DE NORMALIZACIÓN DE TARGETS')
console.log('='.repeat(50) + '\n')

let passed = 0
let failed = 0

tests.forEach((test, i) => {
  const result = normalizeTargetCode(test.input)
  const ok = result === test.expected
  passed += ok ? 1 : 0
  failed += ok ? 0 : 1
  
  const status = ok ? '✅' : '❌'
  console.log(`${status} Test ${i + 1}:`)
  console.log(`   Input:    '${test.input}'`)
  console.log(`   Expected: '${test.expected}'`)
  console.log(`   Result:   '${result}'`)
  if (!ok) console.log(`   ERROR!`)
  console.log()
})

console.log('='.repeat(50))
console.log(`RESULTADO: ${passed}/${tests.length} pasadas`)
console.log('='.repeat(50) + '\n')

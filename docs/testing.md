# Testing Guide

This repo has one rule for product changes:

If behavior changed, the tests should prove it.

## Full Suite

Run the full repo suite before merge:

```bash
npm test
```

Run the capability fixture eval gate when changing contracts, normalization, reviewer-safe extraction behavior, writeback preview/completion, or host capability parity:

```bash
npm run eval:fixtures
```

## High-Signal Focused Suites

### Planner and normalization

```bash
npm test -- services/gateway/tests/requestTemplate.test.ts
```

```bash
npm test -- services/gateway/tests/structuredBody.test.ts
```

```bash
npm test -- services/gateway/tests/fixtureEvalRunner.test.ts
```

### Gateway writeback and execution control

```bash
npm test -- services/gateway/tests/writebackFlow.test.ts
```

```bash
npm test -- services/gateway/tests/executionControl.test.ts
```

```bash
npm test -- services/gateway/tests/hermesClient.test.ts
```

### Excel host

```bash
npm test -- services/gateway/tests/excelWave1Plans.test.ts
```

```bash
npm test -- services/gateway/tests/excelWave2Plans.test.ts
```

```bash
npm test -- services/gateway/tests/excelWave3Plans.test.ts
```

```bash
npm test -- services/gateway/tests/excelWave4Plans.test.ts
```

```bash
npm test -- services/gateway/tests/excelWave5Plans.test.ts
```

```bash
npm test -- services/gateway/tests/excelWave6Plans.test.ts
```

### Google Sheets host

```bash
npm test -- services/gateway/tests/googleSheetsWave1Plans.test.ts
```

```bash
npm test -- services/gateway/tests/googleSheetsWave2Plans.test.ts
```

```bash
npm test -- services/gateway/tests/googleSheetsWave3Plans.test.ts
```

```bash
npm test -- services/gateway/tests/googleSheetsWave4Plans.test.ts
```

```bash
npm test -- services/gateway/tests/googleSheetsWave5Plans.test.ts
```

```bash
npm test -- services/gateway/tests/googleSheetsWave6Plans.test.ts
```

## Which Tests To Run For Which Change

### Contract or response family changes

Run:

- `requestTemplate.test.ts`
- `structuredBody.test.ts`
- `fixtureEvalRunner.test.ts`
- `hermesClient.test.ts`
- `writebackFlow.test.ts`

### Excel-only capability or preview changes

Run:

- the relevant `excelWave*.test.ts`
- `requestTemplate.test.ts` if planner guidance changed
- `writebackFlow.test.ts` if confirm/complete semantics changed

### Google Sheets-only capability or preview changes

Run:

- the relevant `googleSheetsWave*.test.ts`
- `requestTemplate.test.ts` if planner guidance changed
- `writebackFlow.test.ts` if confirm/complete semantics changed

### Execution control, dry-run, undo, redo, or history changes

Run:

- `executionControl.test.ts`
- `executionLedger.test.ts`
- `writebackFlow.test.ts`
- affected host wave 6 tests

## Test Standard For Capability Work

If you add a first-class capability family, cover all of these when relevant:

- planner chooses the family
- model output normalizes into the family
- supported host preview renders it
- supported host executes it exactly
- unsupported host fails closed or stays preview-only
- gateway completion verifies the approved semantics

## Docs-Only Changes

If a change is docs-only:

- no runtime test is required
- say clearly in the PR that the change is docs-only
- avoid claiming runtime verification you did not run

## Before Opening a PR

- run the smallest relevant focused suites while iterating
- run `npm test` before merge
- include the exact commands in the PR description

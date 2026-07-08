// Ambient declaration for exceljs. Its package.json uses the legacy
// `typings` field which doesn't resolve reliably under Node classic
// module resolution. We declare it loosely as `any` so the code compiles;
// runtime is unaffected because exceljs is plain JavaScript underneath.
declare module "exceljs" {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const anyThing: any;
  export default anyThing;
}

export const roundTo = (n, digits = 2) => {
  const factor = 10 ** digits;

  return Math.round(n * factor) / factor;
};

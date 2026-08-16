function commaSeparatedValues(value) {
  return (value ?? '').split(',').filter((entry) => entry.length > 0);
}

export function mutationShardEnvironment(environment) {
  const mutate = commaSeparatedValues(environment.MUTATION_FILES);
  const reportPath = environment.MUTATION_REPORT;
  if (mutate.length === 0) {
    throw new Error('MUTATION_FILES must select at least one source file');
  }
  if (reportPath === undefined || reportPath.length === 0) {
    throw new Error('MUTATION_REPORT must select a JSON report path');
  }
  return {
    excludedMutations: commaSeparatedValues(environment.MUTATION_EXCLUDED),
    mutate,
    reportPath,
  };
}

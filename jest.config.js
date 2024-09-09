module.exports = {
    preset: 'ts-jest',
    testEnvironment: 'node',
    moduleFileExtensions: ['ts', 'js'],
    testRegex: '(/__tests__/.*|(\\.|/)(test|spec))\\.ts$',
    transform: {
      '^.+\\.ts$': 'ts-jest',
    },
    coverageDirectory: 'coverage',
};
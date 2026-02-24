import { test, expect } from 'vitest'
import { makeOutputFilename } from '../src/utils'

test('makeOutputFilename: inserts label before extension', () => {
  expect(makeOutputFilename('data.txt', 'Output')).toBe('dataOutput.txt')
})

test('makeOutputFilename: handles multi-dot filenames', () => {
  expect(makeOutputFilename('data.old.txt', 'Output')).toBe('data.oldOutput.txt')
})

test('makeOutputFilename: handles filename with no extension', () => {
  expect(makeOutputFilename('data', 'Output')).toBe('dataOutput.txt')
})

test('makeOutputFilename: handles null input', () => {
  expect(makeOutputFilename(null, 'Output')).toBe('Output.txt')
})

test('makeOutputFilename: example file with Output label', () => {
  expect(makeOutputFilename('TinyIllustrativeFile.txt', 'Output')).toBe('TinyIllustrativeFileOutput.txt')
})

test('makeOutputFilename: example file with MaxEntOutput label', () => {
  expect(makeOutputFilename('TinyIllustrativeFile.txt', 'MaxEntOutput')).toBe('TinyIllustrativeFileMaxEntOutput.txt')
})

test('makeOutputFilename: example file with FactorialTypology label', () => {
  expect(makeOutputFilename('TinyIllustrativeFile.txt', 'FactorialTypology')).toBe('TinyIllustrativeFileFactorialTypology.txt')
})

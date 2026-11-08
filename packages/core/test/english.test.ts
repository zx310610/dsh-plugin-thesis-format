import { describe, expect, it } from 'vitest'
import { checkKeywordCase } from '../src/rules/english.js'

const fix = (kws: string[], extra?: string[]) =>
  Object.fromEntries(checkKeywordCase(kws, extra).map((i) => [i.keyword, i.suggestion]))

describe('English keyword case', () => {
  it('accepts the template example', () => {
    expect(fix(['traditional industry clusters', 'economy globalization', 'innovation'])).toEqual({})
  })

  it('keeps proper nouns and acronyms (keywords from the sample theses)', () => {
    expect(fix(['Node.js', 'Vue3', 'MySQL', 'Spring Boot', 'React', 'Python', 'WeChat', 'B/S', 'JavaScript', 'IoT'])).toEqual({})
  })

  it('lowercases common words and gives the corrected form', () => {
    expect(fix([
      'Parking Lot Management System',
      'Supermarket management system',
      'WeChat Cloud Development',
      'Front-end and Back-end Separation',
      'Serverless',
      'Game account sharing and rental platform',
      'B/S Architecture',
    ])).toEqual({
      'Parking Lot Management System': 'parking lot management system',
      'Supermarket management system': 'supermarket management system',
      'WeChat Cloud Development': 'WeChat cloud development',
      'Front-end and Back-end Separation': 'front-end and back-end separation',
      Serverless: 'serverless',
      'Game account sharing and rental platform': 'game account sharing and rental platform',
      'B/S Architecture': 'B/S architecture',
    })
  })

  it('does not treat a proper noun inside another word as a match', () => {
    expect(fix(['Gopher Tools'])).toEqual({ 'Gopher Tools': 'gopher tools' })
  })

  it('accepts extra proper nouns from the spec', () => {
    expect(fix(['Gitee Platform'], ['Gitee'])).toEqual({ 'Gitee Platform': 'Gitee platform' })
  })
})

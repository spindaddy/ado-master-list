import { execFile } from 'child_process'
import { tmpdir } from 'os'
import { join } from 'path'
import { writeFileSync } from 'fs'

const SAMPLE_RATE = 44100
const AMP = 32767

function wavFromSamples(samples: Int16Array): Buffer {
  const dataSize = samples.length * 2
  const buf = Buffer.alloc(44 + dataSize)
  buf.write('RIFF', 0)
  buf.writeUInt32LE(36 + dataSize, 4)
  buf.write('WAVE', 8)
  buf.write('fmt ', 12)
  buf.writeUInt32LE(16, 16)
  buf.writeUInt16LE(1, 20)
  buf.writeUInt16LE(1, 22)
  buf.writeUInt32LE(SAMPLE_RATE, 24)
  buf.writeUInt32LE(SAMPLE_RATE * 2, 28)
  buf.writeUInt16LE(2, 32)
  buf.writeUInt16LE(16, 34)
  buf.write('data', 36)
  buf.writeUInt32LE(dataSize, 40)
  for (let i = 0; i < samples.length; i++) {
    buf.writeInt16LE(samples[i], 44 + i * 2)
  }
  return buf
}

function concat(parts: Int16Array[]): Int16Array {
  const total = parts.reduce((n, p) => n + p.length, 0)
  const out = new Int16Array(total)
  let o = 0
  for (const p of parts) {
    out.set(p, o)
    o += p.length
  }
  return out
}

function silence(seconds: number): Int16Array {
  return new Int16Array(Math.floor(SAMPLE_RATE * seconds))
}

function square(freq: number, seconds: number, volume = 1): Int16Array {
  const n = Math.floor(SAMPLE_RATE * seconds)
  const samples = new Int16Array(n)
  const amp = Math.floor(AMP * volume)
  const period = SAMPLE_RATE / freq
  for (let i = 0; i < n; i++) {
    samples[i] = i % period < period / 2 ? amp : -amp
  }
  return samples
}

function siren(seconds = 1.6): Int16Array {
  const n = Math.floor(SAMPLE_RATE * seconds)
  const samples = new Int16Array(n)
  for (let i = 0; i < n; i++) {
    const t = i / SAMPLE_RATE
    const sweep = 0.5 + 0.5 * Math.sin(2 * Math.PI * 2.2 * t)
    const freq = 620 + 780 * sweep
    const period = SAMPLE_RATE / freq
    samples[i] = i % period < period / 2 ? AMP : -AMP
  }
  return samples
}

function klaxon(seconds = 1.4): Int16Array {
  const n = Math.floor(SAMPLE_RATE * seconds)
  const samples = new Int16Array(n)
  for (let i = 0; i < n; i++) {
    const t = i / SAMPLE_RATE
    const freq = 980 - 420 * (t / seconds)
    const buzz = Math.sin(2 * Math.PI * freq * t)
    const gate = Math.sin(2 * Math.PI * 8 * t) >= 0 ? 1 : 0
    samples[i] = Math.floor(AMP * 0.95 * buzz * gate)
  }
  return samples
}

function airHorn(seconds = 0.62): Int16Array {
  const n = Math.floor(SAMPLE_RATE * seconds)
  const samples = new Int16Array(n)
  for (let i = 0; i < n; i++) {
    const t = i / SAMPLE_RATE
    const a = i % (SAMPLE_RATE / 165) < SAMPLE_RATE / 330 ? 1 : -1
    const b = i % (SAMPLE_RATE / 220) < SAMPLE_RATE / 440 ? 1 : -1
    const envelope = Math.min(1, t / 0.02, (seconds - t) / 0.05)
    samples[i] = Math.floor(AMP * 0.98 * envelope * (0.62 * a + 0.38 * b))
  }
  return samples
}

function twoTone(seconds = 2.8): Int16Array {
  const n = Math.floor(SAMPLE_RATE * seconds)
  const samples = new Int16Array(n)
  const slice = 0.32
  for (let i = 0; i < n; i++) {
    const t = i / SAMPLE_RATE
    const freq = Math.floor(t / slice) % 2 === 0 ? 494 : 698
    const period = SAMPLE_RATE / freq
    samples[i] = i % period < period / 2 ? AMP : -AMP
  }
  return samples
}

function phoneNag(): Int16Array {
  const ring = (seconds: number) => {
    const n = Math.floor(SAMPLE_RATE * seconds)
    const samples = new Int16Array(n)
    for (let i = 0; i < n; i++) {
      const t = i / SAMPLE_RATE
      const mix =
        Math.sin(2 * Math.PI * 440 * t) * 0.55 + Math.sin(2 * Math.PI * 480 * t) * 0.45
      const trem = 0.55 + 0.45 * Math.sin(2 * Math.PI * 20 * t)
      samples[i] = Math.floor(AMP * mix * trem)
    }
    return samples
  }
  return concat([ring(1.05), silence(0.18), ring(1.05), silence(0.18), ring(0.85)])
}

function strobeBeep(): Int16Array {
  const parts: Int16Array[] = []
  for (let i = 0; i < 16; i++) {
    parts.push(square(i % 2 === 0 ? 1760 : 1320, 0.09))
    parts.push(silence(0.045))
  }
  return concat(parts)
}

const CUSTOM_WAV: Record<string, () => Int16Array> = {
  Alarm: () => concat([siren(1.05), silence(0.06), siren(1.05), silence(0.06), siren(0.9)]),
  TripleBeep: () =>
    concat([
      square(880, 0.22),
      silence(0.07),
      square(1100, 0.22),
      silence(0.07),
      square(1480, 0.38),
      silence(0.12),
      square(1480, 0.38)
    ]),
  Klaxon: () => concat([klaxon(0.85), silence(0.05), klaxon(0.85), silence(0.05), klaxon(0.7)]),
  RedAlert: () =>
    concat([siren(0.95), silence(0.04), siren(0.95), silence(0.04), siren(0.95), silence(0.04), siren(0.8)]),
  AirHorn: () => concat([airHorn(0.7), silence(0.12), airHorn(0.7), silence(0.12), airHorn(0.85)]),
  PhoneNag: () => phoneNag(),
  StrobeBeep: () => strobeBeep(),
  TwoTone: () => twoTone(3.1)
}

function wavPath(name: string): string {
  const gen = CUSTOM_WAV[name]
  if (!gen) return ''
  const file = join(tmpdir(), `ado-master-alert-${name}.wav`)
  writeFileSync(file, wavFromSamples(gen()))
  return file
}

function afplay(file: string, onDone?: () => void): void {
  execFile('/usr/bin/afplay', ['-v', '1', file], () => onDone?.())
}

function blast(file: string): void {
  afplay(file)
  setTimeout(() => afplay(file), 70)
  setTimeout(() => afplay(file), 140)
}

function speak(text: string): void {
  execFile('/usr/bin/say', ['-v', 'Samantha', '-r', '190', text], () => undefined)
}

const VOICE_LINES: Record<string, string> = {
  Voice: 'Hey. New Azure DevOps ticket.',
  VoiceAlarm: 'Hey. New Azure DevOps ticket.',
  VoiceMail: 'Hey. New Outlook mail.',
  VoiceMailAlarm: 'Hey. New Outlook mail.',
  VoiceTeams: 'Hey. New Teams message.',
  VoiceTeamsAlarm: 'Hey. New Teams message.'
}

const VOICE_WAV: Record<string, string> = {
  VoiceAlarm: 'Alarm',
  VoiceMailAlarm: 'RedAlert',
  VoiceTeamsAlarm: 'AirHorn'
}

const ADHD_LOUD = new Set([
  'RedAlert',
  'AirHorn',
  'PhoneNag',
  'StrobeBeep',
  'TwoTone',
  'VoiceMailAlarm',
  'VoiceTeamsAlarm'
])

export function playNamedAlert(name: string): void {
  if (process.platform !== 'darwin' || name === 'none') return

  const line = VOICE_LINES[name]
  if (line) speak(line)

  const wavName = VOICE_WAV[name] || (CUSTOM_WAV[name] ? name : '')
  if (wavName) {
    const file = wavPath(wavName)
    if (file) {
      blast(file)
      if (ADHD_LOUD.has(name)) setTimeout(() => blast(file), 900)
    }
    return
  }

  if (line) return

  afplay(`/System/Library/Sounds/${name}.aiff`)
}

import DOMPurify from 'dompurify'
import { useEffect, useState, type ReactElement } from 'react'

const ALLOWED_TAGS = [
  'a',
  'b',
  'br',
  'blockquote',
  'code',
  'div',
  'em',
  'h1',
  'h2',
  'h3',
  'h4',
  'h5',
  'h6',
  'hr',
  'i',
  'img',
  'li',
  'ol',
  'p',
  'pre',
  'span',
  'strong',
  'table',
  'tbody',
  'td',
  'th',
  'thead',
  'tr',
  'ul'
]

const ALLOWED_ATTR = [
  'href',
  'src',
  'alt',
  'title',
  'width',
  'height',
  'target',
  'rel',
  'class',
  'style'
]

export function sanitizeAdoHtml(html: string): string {
  if (!html?.trim()) return ''
  return DOMPurify.sanitize(html, {
    ALLOWED_TAGS,
    ALLOWED_ATTR,
    ALLOW_DATA_ATTR: false
  })
}

function decodeHtmlUrl(url: string): string {
  return url
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
}

function isAdoMediaUrl(url: string): boolean {
  try {
    const host = new URL(url).hostname.toLowerCase()
    return (
      host === 'dev.azure.com' ||
      host.endsWith('.dev.azure.com') ||
      host.endsWith('.visualstudio.com') ||
      host.endsWith('.vsassets.io') ||
      host.endsWith('.vssps.visualstudio.com') ||
      host.endsWith('.azureedge.net')
    )
  } catch {
    return false
  }
}

export function toAdoImgProxy(url: string): string {
  const decoded = decodeHtmlUrl(url)
  if (!/^https?:\/\//i.test(decoded) || !isAdoMediaUrl(decoded)) return decoded
  return `adoimg://local/?u=${encodeURIComponent(decoded)}`
}

/** Make relative ADO attachment URLs absolute so auth + loading works. */
export function absolutizeAdoHtml(html: string, organization: string): string {
  if (!html?.trim() || !organization) return html
  const org = organization.replace(/\s+/g, '')
  const orgLower = org.toLowerCase()
  return html.replace(
    /(src|href)=(["'])([^"']+)\2/gi,
    (full, attr: string, quote: string, rawUrl: string) => {
      const url = decodeHtmlUrl(rawUrl)
      if (/^https?:\/\//i.test(url) || url.startsWith('data:') || url.startsWith('adoimg:')) {
        return `${attr}=${quote}${url}${quote}`
      }
      if (url.startsWith('//')) {
        return `${attr}=${quote}https:${url}${quote}`
      }
      if (!url.startsWith('/')) return full

      const first = url.split('/').filter(Boolean)[0] ?? ''
      let absolute = ''
      if (url.toLowerCase().startsWith(`/${orgLower}/`) || first.toLowerCase() === orgLower) {
        absolute = `https://dev.azure.com${url}`
      } else if (url.startsWith('/_apis/')) {
        absolute = `https://dev.azure.com/${org}${url}`
      } else {
        absolute = `https://dev.azure.com/${org}${url}`
      }
      return `${attr}=${quote}${absolute}${quote}`
    }
  )
}

export function proxyAdoMediaHtml(html: string): string {
  if (!html?.trim()) return html
  return html.replace(
    /(src)=(["'])([^"']+)\2/gi,
    (_full, attr: string, quote: string, rawUrl: string) => {
      return `${attr}=${quote}${toAdoImgProxy(rawUrl)}${quote}`
    }
  )
}

export function htmlToPreview(html: string, maxLen = 160): string {
  const text = html
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/\s+/g, ' ')
    .trim()
  if (text.length <= maxLen) return text
  return `${text.slice(0, maxLen)}…`
}

export function RichHtml({
  html,
  organization,
  className,
  empty = '—'
}: {
  html: string
  organization?: string
  className?: string
  empty?: string
}): ReactElement {
  const [zoomSrc, setZoomSrc] = useState<string | null>(null)
  const prepared = proxyAdoMediaHtml(
    sanitizeAdoHtml(organization ? absolutizeAdoHtml(html, organization) : html)
  )

  useEffect(() => {
    if (!zoomSrc) return
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setZoomSrc(null)
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [zoomSrc])

  if (!prepared) {
    return <div className={className}>{empty}</div>
  }
  return (
    <>
      <div
        className={`rich-html ${className ?? ''}`.trim()}
        dangerouslySetInnerHTML={{ __html: prepared }}
        onClick={(e) => {
          const target = e.target
          if (!(target instanceof HTMLImageElement)) return
          e.preventDefault()
          e.stopPropagation()
          setZoomSrc(target.currentSrc || target.src)
        }}
      />
      {zoomSrc && (
        <div
          className="img-lightbox"
          role="dialog"
          aria-modal="true"
          aria-label="Zoomed image"
          onClick={() => setZoomSrc(null)}
        >
          <img
            src={zoomSrc}
            alt="Zoomed attachment"
            onClick={(e) => e.stopPropagation()}
          />
          <button
            type="button"
            className="img-lightbox-close"
            onClick={() => setZoomSrc(null)}
            aria-label="Close"
          >
            ×
          </button>
        </div>
      )}
    </>
  )
}

// PowerPoint slide show state (from ToMain.PRESENTATION_STATE): true when there is no slide or animation left in that direction
export function isPresentationAtEdge(data: { id?: string; stat?: { position?: number; slides?: number; clickIndex?: number; clickCount?: number } } | null, id: string, next: boolean) {
    const stat = data?.id === id ? data.stat : null
    if (!stat?.slides) return false

    const position = stat.position || 0
    if (next) return position >= stat.slides && (stat.clickIndex || 0) >= (stat.clickCount || 0)
    return position <= 1 && !stat.clickIndex
}

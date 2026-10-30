// PowerPoint slide show state (from ToMain.PRESENTATION_STATE): true when there is no slide or animation left in that direction
export function isPresentationAtEdge(data: { id?: string; stat?: { position?: number; slides?: number; clickIndex?: number; clickCount?: number } } | null, id: string, next: boolean) {
    const stat = data?.id === id ? data.stat : null
    if (!stat?.slides) return false

    const position = stat.position || 0
    if (next) return position >= stat.slides && (stat.clickIndex || 0) >= (stat.clickCount || 0)
    return position <= 1 && !stat.clickIndex
}

export type PresentationPreload = "none" | "next" | "all"
export const DEFAULT_PRESENTATION_PRELOAD: PresentationPreload = "next"

// PowerPoint files of a project to keep open in the background: the next one from the active item, all of them, or none
export function getPreloadPaths(items: { id?: string; type?: string }[], activeIndex: number | undefined, mode: PresentationPreload) {
    const presentations = items.map((item, index) => ({ id: item?.id || "", index })).filter((a, i) => a.id && items[i]?.type === "ppt")

    if (mode === "all") return [...new Set(presentations.map((a) => a.id))]
    if (mode !== "next") return []

    const next = presentations.find((a) => a.index >= (activeIndex ?? 0))
    return next ? [next.id] : []
}

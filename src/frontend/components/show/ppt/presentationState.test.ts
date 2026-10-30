import { describe, expect, it } from "vitest"
import { getPreloadPaths, isPresentationAtEdge } from "./presentationState"

const ID = "C:\\slides\\deck.pptx"
let data: any = {}
const setState = (stat: any, id = ID) => (data = { id, stat })

describe("isPresentationAtEdge", () => {
    it("is not at an edge before the slide show reports its state", () => {
        data = {}
        expect(isPresentationAtEdge(data, ID, true)).toBe(false)
        expect(isPresentationAtEdge(data, ID, false)).toBe(false)
    })

    it("ignores the state of another presentation", () => {
        setState({ position: 3, slides: 3, clickIndex: 0, clickCount: 0 }, "other.pptx")
        expect(isPresentationAtEdge(data, ID, true)).toBe(false)
    })

    it("allows next while the last slide still has animations", () => {
        setState({ position: 3, slides: 3, clickIndex: 1, clickCount: 2 })
        expect(isPresentationAtEdge(data, ID, true)).toBe(false)
    })

    it("stops next after the last animation of the last slide", () => {
        setState({ position: 3, slides: 3, clickIndex: 2, clickCount: 2 })
        expect(isPresentationAtEdge(data, ID, true)).toBe(true)
    })

    it("allows previous on the first slide after an animation", () => {
        setState({ position: 1, slides: 3, clickIndex: 1, clickCount: 2 })
        expect(isPresentationAtEdge(data, ID, false)).toBe(false)
    })

    it("stops previous at the start of the first slide", () => {
        setState({ position: 1, slides: 3, clickIndex: 0, clickCount: 2 })
        expect(isPresentationAtEdge(data, ID, false)).toBe(true)
    })

    it("works without animation counts (macOS controller)", () => {
        setState({ position: 3, slides: 3 })
        expect(isPresentationAtEdge(data, ID, true)).toBe(true)
        setState({ position: 2, slides: 3 })
        expect(isPresentationAtEdge(data, ID, true)).toBe(false)
    })
})

describe("getPreloadPaths", () => {
    const items = [
        { id: "song1", type: "show" },
        { id: "a.pptx", type: "ppt" },
        { id: "song2", type: "show" },
        { id: "b.pptx", type: "ppt" },
        { id: "a.pptx", type: "ppt" }
    ]

    it("preloads nothing when turned off", () => {
        expect(getPreloadPaths(items, 0, "none")).toEqual([])
    })

    it("preloads the next PowerPoint from the active item", () => {
        expect(getPreloadPaths(items, 0, "next")).toEqual(["a.pptx"])
        expect(getPreloadPaths(items, 2, "next")).toEqual(["b.pptx"])
    })

    it("keeps the active PowerPoint preloaded", () => {
        expect(getPreloadPaths(items, 3, "next")).toEqual(["b.pptx"])
    })

    it("starts from the top when nothing in the project is active", () => {
        expect(getPreloadPaths(items, undefined, "next")).toEqual(["a.pptx"])
    })

    it("preloads nothing after the last PowerPoint", () => {
        expect(getPreloadPaths(items.slice(0, 3), 2, "next")).toEqual([])
    })

    it("preloads every PowerPoint once", () => {
        expect(getPreloadPaths(items, 0, "all")).toEqual(["a.pptx", "b.pptx"])
    })
})

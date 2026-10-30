import { describe, expect, it } from "vitest"
import { isPresentationAtEdge } from "./presentationState"

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

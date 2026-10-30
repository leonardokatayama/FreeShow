import fs from "fs"
import os from "os"
import path from "path"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

const files = vi.hoisted(() => ({ library: "", picked: [] as string[] }))
vi.mock("../../utils/files", () => ({
    getDataFolderPath: () => {
        fs.mkdirSync(files.library, { recursive: true })
        return files.library
    },
    selectFilesDialog: async () => files.picked
}))

const { addPresentationsToLibrary, getFreeFileName } = await import("./pptLibrary")

const folder = path.join("library")
const existsIn = (names: string[]) => (filePath: string) => names.includes(path.basename(filePath))

describe("getFreeFileName", () => {
    it("keeps the name when it is free", () => {
        expect(getFreeFileName(folder, "Song.pptx", existsIn([]))).toBe("Song.pptx")
    })

    it("adds (2) when the name is taken", () => {
        expect(getFreeFileName(folder, "Song.pptx", existsIn(["Song.pptx"]))).toBe("Song (2).pptx")
    })

    it("counts up to the first free number", () => {
        expect(getFreeFileName(folder, "Song.pptx", existsIn(["Song.pptx", "Song (2).pptx", "Song (3).pptx"]))).toBe("Song (4).pptx")
    })

    it("keeps dots in the name and the extension", () => {
        expect(getFreeFileName(folder, "Culto 30.10.key", existsIn(["Culto 30.10.key"]))).toBe("Culto 30.10 (2).key")
    })
})

describe("addPresentationsToLibrary", () => {
    let root = ""
    let source = ""

    beforeEach(() => {
        root = fs.mkdtempSync(path.join(os.tmpdir(), "freeshow-ppt-library-"))
        source = path.join(root, "source")
        fs.mkdirSync(source)
        files.library = path.join(root, "PowerPoint")
        files.picked = []
    })
    afterEach(() => fs.rmSync(root, { recursive: true, force: true }))

    const pick = (name: string, content = name) => {
        const filePath = path.join(source, name)
        fs.writeFileSync(filePath, content)
        files.picked.push(filePath)
        return filePath
    }

    it("copies the picked files into the library and returns their new paths", async () => {
        const original = pick("Culto.pptx", "deck")

        const added = await addPresentationsToLibrary()

        expect(added).toEqual([path.join(files.library, "Culto.pptx")])
        expect(fs.readFileSync(added[0], "utf8")).toBe("deck")
        // the original stays where it was
        expect(fs.existsSync(original)).toBe(true)
    })

    it("adds (2) to a copy whose name is already in the library", async () => {
        fs.mkdirSync(files.library, { recursive: true })
        fs.writeFileSync(path.join(files.library, "Culto.pptx"), "old")
        pick("Culto.pptx", "new")

        const added = await addPresentationsToLibrary()

        expect(added).toEqual([path.join(files.library, "Culto (2).pptx")])
        expect(fs.readFileSync(path.join(files.library, "Culto.pptx"), "utf8")).toBe("old")
        expect(fs.readFileSync(added[0], "utf8")).toBe("new")
    })

    it("changes nothing when the file picker is cancelled", async () => {
        const added = await addPresentationsToLibrary()

        expect(added).toEqual([])
        expect(fs.existsSync(files.library)).toBe(false)
    })

    it("keeps a file that is already in the library without copying it again", async () => {
        fs.mkdirSync(files.library, { recursive: true })
        const inLibrary = path.join(files.library, "Culto.pptx")
        fs.writeFileSync(inLibrary, "deck")
        files.picked = [inLibrary]

        const added = await addPresentationsToLibrary()

        expect(added).toEqual([inLibrary])
        expect(fs.readdirSync(files.library)).toEqual(["Culto.pptx"])
    })

    it("skips a file that can't be copied and keeps the others", async () => {
        files.picked.push(path.join(source, "missing.pptx"))
        pick("Culto.pptx")
        vi.spyOn(console, "error").mockImplementation(() => {})

        const added = await addPresentationsToLibrary()

        expect(added).toEqual([path.join(files.library, "Culto.pptx")])
    })
})

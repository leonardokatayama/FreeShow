import path from "path"
import { describe, expect, it, vi } from "vitest"

vi.mock("../../utils/files", () => ({ getDataFolderPath: vi.fn(), selectFilesDialog: vi.fn() }))

const { getFreeFileName } = await import("./pptLibrary")

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

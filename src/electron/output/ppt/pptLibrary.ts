// PowerPoint library: Documents/FreeShow/PowerPoint, where added presentations are copied to (PowerPoints drawer tab)

import fs from "fs"
import path from "path"
import { getDataFolderPath, selectFilesDialog } from "../../utils/files"

const PRESENTATION_EXTENSIONS = ["ppt", "pptx", "key"]

export function getPresentationLibraryPath() {
    return getDataFolderPath("powerpoint")
}

// "Song.pptx" -> "Song (2).pptx", "Song (3).pptx"... while the name is taken
export function getFreeFileName(folder: string, fileName: string, exists = (filePath: string) => fs.existsSync(filePath)) {
    const extension = path.extname(fileName)
    const baseName = path.basename(fileName, extension)

    let name = fileName
    for (let count = 2; exists(path.join(folder, name)); count++) name = `${baseName} (${count})${extension}`
    return name
}

// pick presentations and copy them into the library, returns their library paths
export async function addPresentationsToLibrary() {
    const files = await selectFilesDialog("", { name: "PowerPoint", extensions: PRESENTATION_EXTENSIONS })
    if (!files.length) return []

    const folder = getPresentationLibraryPath()
    const added: string[] = []
    for (const file of files) {
        // already in the library
        if (path.resolve(path.dirname(file)) === path.resolve(folder)) {
            added.push(file)
            continue
        }

        const destination = path.join(folder, getFreeFileName(folder, path.basename(file)))
        try {
            await fs.promises.copyFile(file, destination)
            added.push(destination)
        } catch (err) {
            console.error("Could not add presentation to library:", file, err)
        }
    }

    return added
}

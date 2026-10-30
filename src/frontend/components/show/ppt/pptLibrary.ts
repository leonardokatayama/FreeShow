import { Main } from "../../../../types/IPC/Main"
import { addToProject } from "../../../converters/project"
import { requestMain } from "../../../IPC/main"

// pick PowerPoint files and copy them into Documents/FreeShow/PowerPoint, returns their library paths
export async function addPresentationsToLibrary(): Promise<string[]> {
    return (await requestMain(Main.PRESENTATION_ADD)) || []
}

// Import > PowerPoint: copy into the library and add to the current project as PowerPoint items (with animations)
export async function importPresentationsToProject() {
    const paths = await addPresentationsToLibrary()
    if (paths.length) addToProject("ppt", paths)
}

import { Main } from "../../../../types/IPC/Main"
import { requestMain } from "../../../IPC/main"

// pick PowerPoint files and copy them into Documents/FreeShow/PowerPoint, returns their library paths
export async function addPresentationsToLibrary(): Promise<string[]> {
    return (await requestMain(Main.PRESENTATION_ADD)) || []
}

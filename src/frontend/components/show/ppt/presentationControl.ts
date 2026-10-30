import { get } from "svelte/store"
import { Main } from "../../../../types/IPC/Main"
import type { ToMain, ToMainSendPayloads } from "../../../../types/IPC/ToMain"
import { sendMain } from "../../../IPC/main"
import { outLocked, outputs, presentationData, special } from "../../../stores"
import { getFileName, removeExtension } from "../../helpers/media"
import { getActiveOutputs, setOutput } from "../../helpers/output"
import { clearBackground, clearSlide } from "../../output/clear"

// PowerPoint/Keynote slide show started from the PowerPoint view or the API (Remote)

let starting: { path: string; name: string } | null = null

export function startPresentation(path: string, name = "") {
    if (!path || get(outLocked)) return

    starting = { path, name: name || removeExtension(getFileName(path)) }
    const outputId = getActiveOutputs(get(outputs), false, true, true)[0]
    sendMain(Main.START_SLIDESHOW, { path, program: get(special).presentationApp || "PowerPoint", outputId })
    clearSlide()
}

export function gotoPresentationSlide(index: number) {
    if (get(outLocked)) return
    sendMain(Main.PRESENTATION_CONTROL, { action: "goto", index })
}

export function receivePresentationState(data: ToMainSendPayloads[ToMain.PRESENTATION_STATE]) {
    presentationData.set(data)

    const isOutputted = get(outputs)[getActiveOutputs()[0]]?.out?.slide?.id === data.id
    // native PowerPoint slide show ended (e.g. Esc pressed): clear it from the outputs
    if (data.ended) {
        if (starting?.path === data.id) starting = null
        if (isOutputted) clearSlide()
        return
    }

    // the slide show is running: show it on the outputs
    if (starting?.path !== data.id) return
    setOutput("slide", { type: "ppt", id: data.id, name: starting.name, page: (data.stat?.position || 1) - 1, pages: data.stat?.slides })
    clearBackground()
    starting = null
}

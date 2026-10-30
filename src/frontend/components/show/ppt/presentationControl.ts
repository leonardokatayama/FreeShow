import { get } from "svelte/store"
import { Main } from "../../../../types/IPC/Main"
import { REMOTE } from "../../../../types/Channels"
import type { ToMain, ToMainSendPayloads } from "../../../../types/IPC/ToMain"
import { sendMain } from "../../../IPC/main"
import { activeProject, activeShow, os, outLocked, outputs, presentationData, projects, special } from "../../../stores"
import { sendData } from "../../../utils/sendData"
import { getFileName, removeExtension } from "../../helpers/media"
import { getActiveOutputs, setOutput } from "../../helpers/output"
import { clearBackground, clearSlide } from "../../output/clear"
import { DEFAULT_PRESENTATION_PRELOAD, getPreloadPaths } from "./presentationState"

// PowerPoint/Keynote slide show started from the PowerPoint view or the API (Remote)

let starting: { path: string; name: string } | null = null

export function startPresentation(path: string, name = "") {
    if (!path || get(outLocked)) return

    // clear first: stopping a running slide show must reach the app before the new one starts
    clearSlide()
    starting = { path, name: name || removeExtension(getFileName(path)) }
    const outputId = getActiveOutputs(get(outputs), false, true, true)[0]
    sendMain(Main.START_SLIDESHOW, { path, program: get(special).presentationApp || "PowerPoint", outputId })
}

export function gotoPresentationSlide(index: number) {
    if (get(outLocked)) return
    sendMain(Main.PRESENTATION_CONTROL, { action: "goto", index })
}

export function receivePresentationState(data: ToMainSendPayloads[ToMain.PRESENTATION_STATE]) {
    presentationData.set(data)
    sendData(REMOTE, { channel: "OUT_DATA" })

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

// keep the project's PowerPoint files open in the background (Settings > Outputs), so they start instantly
let preloaded = ""
export function watchPresentationPreload() {
    const update = () => {
        if (get(os).platform !== "win32") return

        const items = get(projects)[get(activeProject) || ""]?.shows || []
        const paths = getPreloadPaths(items, get(activeShow)?.index, get(special).presentationPreload || DEFAULT_PRESENTATION_PRELOAD)

        const key = JSON.stringify(paths)
        if (key === preloaded) return
        preloaded = key
        sendMain(Main.PRESENTATION_PRELOAD, { paths })
    }

    ;[os, special, projects, activeProject, activeShow].forEach((store) => store.subscribe(update))
}

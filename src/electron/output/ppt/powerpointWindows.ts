// Native PowerPoint control on Windows.
// PowerPoint runs hidden (no editing window, no taskbar icon) and only its full screen
// slide show window is placed over the active output, so animations, transitions and
// embedded media play natively. FreeShow acts as the remote control.

import { type ChildProcessWithoutNullStreams, spawn } from "child_process"
import crypto from "crypto"
import { app, screen } from "electron"
import fs from "fs"
import path from "path"
import { getMainWindow } from "../.."
import { ToMain } from "../../../types/IPC/ToMain"
import { sendToMain } from "../../IPC/main"
import { createFolder, getDataFolderPath, sanitizeFileName } from "../../utils/files"
import { OutputHelper } from "../OutputHelper"
import { POWERPOINT_SCRIPT } from "./powerpointScript"

type Bounds = { x: number; y: number; width: number; height: number }
type PptState = { running: boolean; position?: number; slides?: number; state?: number; clickIndex?: number; clickCount?: number }
type PptInfo = { titles: string[]; notes: string[] }

const THUMBNAIL_WIDTH = 480
const STATE_POLL_INTERVAL = 400
const OPEN_TIMEOUT = 60000
const ACTION_TIMEOUT = 15000

// ----- BRIDGE -----

class PowerPointBridge {
    private process: ChildProcessWithoutNullStreams | null = null
    private buffer = ""
    private requestId = 0
    private pending = new Map<number, { resolve: (value: any) => void; reject: (reason: any) => void; timer: NodeJS.Timeout }>()
    onExit: (() => void) | null = null

    get busy() {
        return this.pending.size > 0
    }

    private start() {
        if (this.process) return

        const scriptPath = path.join(app.getPath("temp"), "freeshow-powerpoint.ps1")
        fs.writeFileSync(scriptPath, POWERPOINT_SCRIPT, "utf8")

        const child = spawn("powershell.exe", ["-NoLogo", "-NoProfile", "-NonInteractive", "-Sta", "-ExecutionPolicy", "Bypass", "-File", scriptPath], { windowsHide: true })
        this.process = child

        child.stdout.setEncoding("utf8")
        child.stdout.on("data", (chunk: string) => this.receive(chunk))
        child.stderr.setEncoding("utf8")
        child.stderr.on("data", (chunk: string) => console.error("PowerPoint bridge:", chunk))

        child.on("exit", () => {
            if (this.process !== child) return
            this.process = null
            this.buffer = ""
            this.pending.forEach(({ reject, timer }) => {
                clearTimeout(timer)
                reject(new Error("PowerPoint bridge exited"))
            })
            this.pending.clear()
            this.onExit?.()
        })
        child.on("error", (err) => console.error("PowerPoint bridge error:", err))
    }

    private receive(chunk: string) {
        this.buffer += chunk
        let newLine = this.buffer.indexOf("\n")
        while (newLine >= 0) {
            const line = this.buffer.slice(0, newLine).trim()
            this.buffer = this.buffer.slice(newLine + 1)
            newLine = this.buffer.indexOf("\n")
            if (!line) continue

            let response: { id: number | null; ok: boolean; result?: any; error?: string }
            try {
                response = JSON.parse(line)
            } catch {
                console.error("PowerPoint bridge: invalid response", line)
                continue
            }

            const request = response.id === null ? null : this.pending.get(response.id)
            if (!request) continue
            this.pending.delete(response.id!)
            clearTimeout(request.timer)

            if (response.ok) request.resolve(response.result)
            else request.reject(new Error(response.error || "Unknown PowerPoint error"))
        }
    }

    request<T = any>(action: string, data: any = {}, timeout = ACTION_TIMEOUT): Promise<T> {
        this.start()

        const id = ++this.requestId
        return new Promise<T>((resolve, reject) => {
            const timer = setTimeout(() => {
                this.pending.delete(id)
                reject(new Error(`PowerPoint did not respond to "${action}"`))
            }, timeout)

            this.pending.set(id, { resolve, reject, timer })
            this.process!.stdin.write(JSON.stringify({ id, action, data }) + "\n", "utf8")
        })
    }

    stop() {
        if (!this.process) return
        // closing stdin makes the script close the presentation and exit by itself
        const child = this.process
        child.stdin.end()
        // the next request starts a new process instead of writing to the closed one
        this.process = null
        this.buffer = ""
        this.pending.forEach(({ reject, timer }) => {
            clearTimeout(timer)
            reject(new Error("PowerPoint bridge stopped"))
        })
        this.pending.clear()
        setTimeout(() => {
            if (child.exitCode === null) child.kill()
        }, 10000)
    }
}

const bridge = new PowerPointBridge()

// ----- SESSION -----

let session = 0
let currentPath = ""
let info: PptInfo = { titles: [], notes: [] }
let thumbnails: string[] = []
let lastState: PptState = { running: false }
let pollTimer: NodeJS.Timeout | null = null
let onEndedCallback: (() => void) | null = null

bridge.onExit = () => {
    if (currentPath) ended()
}

export function isWindowsSlideshowActive() {
    return !!currentPath
}

export async function startWindowsSlideshow(filePath: string, outputId: string | undefined, onEnded: () => void) {
    stopPolling()
    const id = ++session
    currentPath = filePath
    onEndedCallback = onEnded
    info = { titles: [], notes: [] }
    thumbnails = []
    lastState = { running: false }

    let result: { state: PptState; info: PptInfo }
    try {
        result = await bridge.request("open", { path: filePath, bounds: getTargetBounds(outputId) }, OPEN_TIMEOUT)
    } catch (err) {
        if (id !== session) return false
        console.error("PowerPoint OPEN", err)
        sendToMain(ToMain.ALERT, "Could not start PowerPoint. Make sure Microsoft PowerPoint is installed.<br>" + (err as Error).message)
        currentPath = ""
        onEndedCallback = null
        return false
    }
    if (id !== session) return false

    info = { titles: result.info?.titles || [], notes: result.info?.notes || [] }
    lastState = result.state
    sendState()

    // keep keyboard control in FreeShow
    getMainWindow()?.focus()

    startPolling(id)
    generateThumbnails(id, filePath, result.state.slides || 0)
    return true
}

export async function windowsPresentationControl(action: string, index?: number) {
    if (!currentPath) return

    if (action === "stop") {
        await stopWindowsSlideshow()
        return
    }

    try {
        const state = await bridge.request<PptState>(action, { index })
        updateState(state)
    } catch (err) {
        console.error("PowerPoint action failed:", action, err)
    }
}

export async function stopWindowsSlideshow() {
    if (!currentPath) return
    const id = ++session
    stopPolling()

    try {
        await bridge.request("close")
    } catch (err) {
        console.error("PowerPoint CLOSE", err)
    }

    // a new slide show started while this one was closing: the bridge is already running it
    if (id !== session) return

    bridge.stop()
    ended()
}

function ended() {
    stopPolling()
    const filePath = currentPath
    currentPath = ""
    lastState = { running: false }
    if (filePath) sendToMain(ToMain.PRESENTATION_STATE, { id: filePath, native: true, ended: true, stat: lastState, info, thumbnails })

    const callback = onEndedCallback
    onEndedCallback = null
    callback?.()
}

// ----- STATE -----

function startPolling(id: number) {
    stopPolling()
    pollTimer = setInterval(async () => {
        if (id !== session) return stopPolling()
        // don't queue polls behind navigation or thumbnail exports
        if (bridge.busy) return

        try {
            const state = await bridge.request<PptState>("state")
            if (id === session) updateState(state)
        } catch {
            // PowerPoint can reject calls while busy rendering, try again on next poll
        }
    }, STATE_POLL_INTERVAL)
}

function stopPolling() {
    if (pollTimer) clearInterval(pollTimer)
    pollTimer = null
}

function updateState(state: PptState) {
    if (!currentPath) return

    // slide show was closed from PowerPoint (e.g. Esc)
    if (!state.running) {
        stopWindowsSlideshow()
        return
    }

    const changed = (Object.keys(state) as (keyof PptState)[]).some((key) => state[key] !== lastState[key])
    lastState = state
    if (changed) sendState()
}

function sendState() {
    sendToMain(ToMain.PRESENTATION_STATE, { id: currentPath, native: true, stat: lastState, info, thumbnails })
}

// ----- THUMBNAILS -----

async function generateThumbnails(id: number, filePath: string, slides: number) {
    if (!slides) return

    const folder = getThumbnailFolder(filePath)
    thumbnails = new Array(slides).fill("")

    for (let index = 1; index <= slides; index++) {
        if (id !== session) return

        const thumbnailPath = path.join(folder, `${index}.png`)
        if (!fs.existsSync(thumbnailPath)) {
            try {
                await bridge.request("thumbnail", { index, path: thumbnailPath, width: THUMBNAIL_WIDTH }, 30000)
            } catch (err) {
                console.error("PowerPoint thumbnail", index, err)
                continue
            }
        }

        if (id !== session) return
        thumbnails[index - 1] = thumbnailPath
        // update in batches to not flood the UI
        if (index % 5 === 0 || index === slides) sendState()
    }
}

// cached per file version
function getThumbnailFolder(filePath: string) {
    let stat = { mtimeMs: 0, size: 0 }
    try {
        stat = fs.statSync(filePath)
    } catch {
        // file missing, PowerPoint will report the error
    }

    const hash = crypto.createHash("md5").update(`${filePath}|${stat.mtimeMs}|${stat.size}`).digest("hex").slice(0, 10)
    const name = sanitizeFileName(path.basename(filePath, path.extname(filePath)))
    return createFolder(path.join(getDataFolderPath("imports", "PowerPoint"), "Slides", `${name}_${hash}`))
}

// ----- OUTPUT -----

// physical pixel bounds of the output the slide show should cover
function getTargetBounds(outputId?: string): Bounds | null {
    const outputs = OutputHelper.getAllOutputs().filter((output) => output.window && !output.window.isDestroyed() && !output.invisible && output.window.isVisible())
    const output = (outputId ? outputs.find((a) => a.id === outputId) : null) || outputs[0]
    if (!output) return null

    return screen.dipToScreenRect(output.window, output.window.getBounds())
}

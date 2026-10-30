<script lang="ts">
    import { onDestroy } from "svelte"
    import { uid } from "uid"
    import { Main } from "../../../../types/IPC/Main"
    import { ToMain } from "../../../../types/IPC/ToMain"
    import type { FileFolder } from "../../../../types/Main"
    import { destroyMain, receiveToMain, requestMain, sendMain } from "../../../IPC/main"
    import { drawerTabsData, labelsDisabled, pptFolders } from "../../../stores"
    import { getAccess } from "../../../utils/profile"
    import { keysToID, sortObject } from "../../helpers/array"
    import { addDrawerFolder } from "../../helpers/dropActions"
    import Icon from "../../helpers/Icon.svelte"
    import { getExtension, getMediaType } from "../../helpers/media"
    import T from "../../helpers/T.svelte"
    import MaterialButton from "../../inputs/MaterialButton.svelte"
    import NavigationSections from "./NavigationSections.svelte"

    const profile = getAccess("powerpoint")
    $: readOnly = profile.global === "read"

    $: activeSubTab = $drawerTabsData.powerpoint?.activeSubTab || ""

    // Documents/FreeShow/PowerPoint: always listed, can't be removed
    let libraryPath = ""
    requestMain(Main.PRESENTATION_LIBRARY, undefined, (path) => (libraryPath = path || ""))

    $: foldersList = keysToID($pptFolders)

    let counts: { [path: string]: number } = {}
    $: if (libraryPath || foldersList.length) getCounts()
    async function getCounts() {
        const folderPaths = [libraryPath, ...foldersList.map((a) => a.path || "")].filter(Boolean)
        const data: FileFolder[] = Object.values((await requestMain(Main.READ_FOLDER, { path: folderPaths })) || {})

        const newCounts: { [path: string]: number } = {}
        folderPaths.forEach((folderPath) => {
            const files: string[] = (data.find((a) => a.path === folderPath) as any)?.files || []
            newCounts[folderPath] = files.filter((filePath) => getMediaType(getExtension(filePath)) === "ppt").length
        })
        counts = newCounts
    }

    $: allCount = Object.values(counts).reduce((total, count) => total + count, 0)

    let sections: any[] = []
    $: sections = [
        [
            { id: "all", label: "category.all", icon: "all", count: allCount },
            { id: "library", label: "tabs.powerpoint_library", icon: "folder", count: counts[libraryPath] || 0 }
        ],
        [{ id: "TITLE", label: "media.folders" }, ...sortObject(foldersList, "name").map((a) => ({ id: a.id, label: a.name, icon: a.icon || "folder", count: counts[a.path || ""] || 0, boxedIcon: true }))]
    ]

    const PICK_ID = uid()
    function addFolder() {
        sendMain(Main.OPEN_FOLDER, { channel: PICK_ID })
    }
    let listenerId = receiveToMain(ToMain.OPEN_FOLDER2, (data) => {
        if (data.channel !== PICK_ID || !data.path) return
        addDrawerFolder(data, "powerpoint")
    })
    onDestroy(() => destroyMain(listenerId))

    function updateName(e: any) {
        const { id, value } = e.detail
        pptFolders.update((a) => {
            if (!a[id]) return a
            a[id].name = value
            return a
        })
    }
</script>

<NavigationSections {sections} active={activeSubTab} on:rename={updateName}>
    <div slot="section_1" style="padding: 8px;{foldersList.length ? 'padding-top: 12px;' : ''}">
        <MaterialButton style="width: 100%;" title="new.system_folder" variant="outlined" disabled={readOnly} on:click={addFolder} small>
            <Icon id="add" size={$labelsDisabled ? 0.9 : 1} white={$labelsDisabled} />
            {#if !$labelsDisabled}<T id="new.system_folder" />{/if}
        </MaterialButton>
    </div>
</NavigationSections>

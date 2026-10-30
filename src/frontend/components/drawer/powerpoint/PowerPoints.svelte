<script lang="ts">
    import { Main } from "../../../../types/IPC/Main"
    import type { FileFolder } from "../../../../types/Main"
    import { requestMain } from "../../../IPC/main"
    import { addProjectItem } from "../../../converters/project"
    import { activeShow, labelsDisabled, outLocked, pptFolders } from "../../../stores"
    import { getAccess } from "../../../utils/profile"
    import { sortFilenames } from "../../helpers/array"
    import Icon from "../../helpers/Icon.svelte"
    import { getExtension, getMediaType, removeExtension } from "../../helpers/media"
    import T from "../../helpers/T.svelte"
    import FloatingInputs from "../../input/FloatingInputs.svelte"
    import Button from "../../inputs/Button.svelte"
    import MaterialButton from "../../inputs/MaterialButton.svelte"
    import Center from "../../system/Center.svelte"
    import SelectElem from "../../system/SelectElem.svelte"
    import { addPresentationsToLibrary } from "../../show/ppt/pptLibrary"

    export let active: string | null
    export let searchValue = ""

    const readOnly = getAccess("powerpoint").global === "read"

    let libraryPath = ""
    requestMain(Main.PRESENTATION_LIBRARY, undefined, (path) => (libraryPath = path || ""))

    // "all": library + every folder, "library": library, otherwise one of the added system folders
    $: folderPaths = !libraryPath ? [] : active === "library" ? [libraryPath] : active && $pptFolders[active] ? [$pptFolders[active].path || ""] : [libraryPath, ...Object.values($pptFolders).map((a) => a.path || "")]
    $: requestFiles(folderPaths)

    let files: FileFolder[] = []
    let requesting = 0
    async function requestFiles(paths: string[]) {
        const request = ++requesting
        if (!paths.filter(Boolean).length) {
            files = []
            return
        }

        const data = await requestMain(Main.READ_FOLDER, { path: paths.filter(Boolean) })
        if (request !== requesting) return
        files = sortFilenames(Object.values(data || {}).filter((a) => !a.isFolder && getMediaType(getExtension(a.name)) === "ppt"))
    }

    $: searchedFiles = searchValue ? files.filter((a) => a.name.toLowerCase().includes(searchValue.toLowerCase())) : files

    function open(file: FileFolder) {
        activeShow.set({ id: file.path, type: "ppt", name: removeExtension(file.name) })
    }

    async function addFiles() {
        const added = await addPresentationsToLibrary()
        if (added.length) requestFiles(folderPaths)
    }
</script>

<div class="scroll">
    {#if searchedFiles.length}
        <div class="list">
            {#each searchedFiles as file (file.path)}
                <SelectElem id="powerpoint" data={{ path: file.path, name: file.name }} draggable>
                    <Button
                        active={$activeShow?.id === file.path}
                        border
                        style="width: 100%;"
                        title={file.path}
                        bold={false}
                        on:click={(e) => {
                            if ($outLocked || e.ctrlKey || e.metaKey || e.shiftKey) return
                            open(file)
                        }}
                        on:dblclick={(e) => {
                            if (e.ctrlKey || e.metaKey || e.shiftKey) return
                            addProjectItem({ id: file.path, name: removeExtension(file.name), type: "ppt" })
                        }}
                    >
                        <Icon id="ppt" box={50} right />
                        <p>{removeExtension(file.name)}</p>
                    </Button>
                </SelectElem>
            {/each}
        </div>
    {:else}
        <Center faded>
            <T id="empty.general" />
        </Center>
    {/if}
</div>

<FloatingInputs onlyOne>
    <MaterialButton icon="add" style="flex: 1;" disabled={readOnly} on:click={addFiles} center title="tabs.powerpoint_add">
        {#if !$labelsDisabled}<T id="tabs.powerpoint_add" />{/if}
    </MaterialButton>
</FloatingInputs>

<style>
    .scroll {
        flex: 1;
        height: 100%;
        overflow-y: auto;
    }

    .list {
        display: flex;
        flex-direction: column;
        padding: 5px;
        /* room for the floating add button */
        padding-bottom: 60px;
    }

    .list p {
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
    }
</style>

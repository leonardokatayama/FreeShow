<script lang="ts">
    import { autoOutput, os, special } from "../../../stores"
    import { translateText } from "../../../utils/language"
    import MaterialDropdown from "../../inputs/MaterialDropdown.svelte"
    import MaterialToggleSwitch from "../../inputs/MaterialToggleSwitch.svelte"
    import { DEFAULT_PRESENTATION_PRELOAD } from "../../show/ppt/presentationState"
    import Tip from "../../main/Tip.svelte"

    function updateSpecial(value: any, key: string, allowEmpty = false) {
        special.update((a) => {
            if (!allowEmpty && !value) delete a[key]
            else a[key] = value

            return a
        })
    }

    $: presentationPreloadOptions = ["none", "next", "all"].map((value) => ({ value, label: translateText(`settings.preload_powerpoint_${value}`) }))
</script>

<Tip type="info" value="tips.global_options" bottom={20} />

<MaterialToggleSwitch label="settings.auto_output" checked={$autoOutput} defaultValue={false} on:change={(e) => autoOutput.set(e.detail)} />
<!-- apparently doesn't work on some versions of macOS -->
{#if $os.platform !== "darwin" || $special.hideCursor}
    <MaterialToggleSwitch label="settings.hide_cursor_in_output" checked={$special.hideCursor} defaultValue={false} on:change={(e) => updateSpecial(e.detail, "hideCursor")} />
{/if}

<!-- PowerPoint is controlled natively on Windows only -->
{#if $os.platform === "win32"}
    <MaterialDropdown label="settings.preload_powerpoint" value={$special.presentationPreload || DEFAULT_PRESENTATION_PRELOAD} defaultValue={DEFAULT_PRESENTATION_PRELOAD} options={presentationPreloadOptions} on:change={(e) => updateSpecial(e.detail, "presentationPreload")} />
{/if}

<script lang="ts">
    import { onDestroy, onMount } from "svelte"
    import { Main } from "../../../../types/IPC/Main"
    import type { ShowRef } from "../../../../types/Projects"
    import { requestMain, sendMain } from "../../../IPC/main"
    import { activeShow, os, outLocked, outputs, presentationApps, presentationData, special } from "../../../stores"
    import { translateText } from "../../../utils/language"
    import Icon from "../../helpers/Icon.svelte"
    import { encodeFilePath } from "../../helpers/media"
    import { getOutputContent } from "../../helpers/output"
    import T from "../../helpers/T.svelte"
    import Button from "../../inputs/Button.svelte"
    import Dropdown from "../../inputs/Dropdown.svelte"
    import Center from "../../system/Center.svelte"
    import { gotoPresentationSlide, startPresentation } from "./presentationControl"
    import ScreenCapture from "./ScreenCapture.svelte"

    export let show: ShowRef

    // WIP use savedScreen = $projects[$activeProject || ""].shows.find((a) => a.id === path)?.data?.screenName to determine if it is active or not

    onMount(() => {
        requestMain(Main.SLIDESHOW_GET_APPS, undefined, (a) => presentationApps.set(a || []))
    })

    let opening = false
    let retry = false
    function newPresentation() {
        if ($outLocked) return

        reset()
        opening = true
        startPresentation(show.id, show.name)

        setTimeout(() => (opening ? (retry = true) : ""), 8000)
    }

    function keydown(e: KeyboardEvent) {
        if ($outLocked || !$presentationApps?.length) return

        // advancing is handled by the global shortcuts (OutputHelper), this only starts the slide show
        if (e.key === " " && outSlide?.id !== show.id && !opening) {
            e.preventDefault()
            newPresentation()
        }
    }

    $: outSlide = getOutputContent("", $outputs)

    // the output is set by receivePresentationState once the slide show runs
    $: if (opening && outSlide?.id === show.id) reset()

    onDestroy(() => {
        if ($activeShow?.type !== "ppt") sendMain(Main.PRESENTATION_CONTROL, { action: "stop" })
    })

    $: if (show.id) reset()
    function reset() {
        opening = false
        retry = false
    }

    function restartPresentation() {
        presentationData.set({})
        newPresentation()
    }

    // native (Windows): PowerPoint runs hidden, its slide show covers the output
    $: native = $presentationData?.id === show.id && !!$presentationData?.native
    $: currentPosition = $presentationData?.stat?.position || 0
    $: thumbnails = ($presentationData?.thumbnails || []) as string[]
    $: slideCount = $presentationData?.stat?.slides || 0

    const gotoSlide = gotoPresentationSlide

    // keep the current slide in view
    let grid: HTMLElement | undefined
    $: if (grid && currentPosition) grid.querySelector(`[data-index="${currentPosition}"]`)?.scrollIntoView({ block: "nearest" })
</script>

<svelte:window on:keydown={keydown} />

{#if $presentationApps === null}
    <Center faded><T id="remote.loading" /></Center>
{:else if !$presentationApps.length}
    <Center>
        <T id="presentation_control.unsupported" />
        {$os.platform || "this OS"}.
        <br />
        <T id="presentation_control.unsupported_tip" />
    </Center>
{:else if outSlide?.id === show.id && native}
    <div class="grid" bind:this={grid}>
        {#each Array(slideCount) as _, i}
            {@const index = i + 1}
            <button class="thumbnail" class:active={index === currentPosition} data-index={index} on:click={() => gotoSlide(index)} title={$presentationData.info?.titles?.[i] || ""}>
                {#if thumbnails[i]}
                    <img src={encodeFilePath(thumbnails[i])} alt="" draggable="false" />
                {/if}
                <span class="number">{index}</span>
            </button>
        {/each}
    </div>

    <div class="info">
        <Button on:click={restartPresentation} title={translateText("presentation_control.restart")}>
            <Icon id="refresh" />
        </Button>
        <p style="white-space: normal;overflow: auto;padding: 3px 8px;">
            <b>{currentPosition}/{slideCount}</b>
            {#if $presentationData.stat?.clickCount}
                <span style="opacity: 0.6;padding-inline-start: 6px;">({$presentationData.stat.clickIndex}/{$presentationData.stat.clickCount})</span>
            {/if}
            <b style="padding-inline-start: 10px;">{$presentationData.info?.titles?.[currentPosition - 1] || ""}</b>
            <span style="padding-inline-start: 10px;">{$presentationData.info?.notes?.[currentPosition - 1] || ""}</span>
        </p>
    </div>
{:else if outSlide?.id === show.id}
    <div class="fill">
        <ScreenCapture path={show.id} />
    </div>

    {#if $presentationData?.id === show.id && $presentationData?.stat?.slides}
        <div class="info">
            <Button on:click={restartPresentation} title={translateText("presentation_control.restart")}>
                <Icon id="refresh" />
            </Button>
            <p style="white-space: normal;overflow: auto;padding: 3px 8px;">
                <b>{$presentationData.info?.titles?.[$presentationData.stat?.position - 1] || ""}</b>
                <span style="padding-inline-start: 10px;">{$presentationData.info?.notes[$presentationData.stat?.position - 1] || ""}</span>
            </p>
        </div>
    {/if}
{:else if opening}
    <Center>
        <p><T id="presentation_control.opening" /></p>
        {#if retry}
            <br />
            <T id="presentation_control.retry" />
            <Button on:click={newPresentation} style="margin-top: 8px;"><T id="presentation_control.try_again" /></Button>
        {/if}
    </Center>
{:else}
    <div style="display: flex;flex-direction: column;height: 100%;">
        <Dropdown
            options={$presentationApps.map((id) => ({ name: id }))}
            value={$special.presentationApp || "PowerPoint"}
            on:click={(e) => {
                special.update((a) => {
                    a.presentationApp = e.detail?.name
                    return a
                })
            }}
        />

        <div class="fill">
            <Button on:click={newPresentation} style="font-size: 4em;width: 100%;" center>
                <Icon id="play" size={6} right />
                <T id="presentation_control.start" />
            </Button>
        </div>
    </div>
{/if}

<style>
    .fill {
        position: relative;
        height: 100%;

        display: flex;
        /* align-items: center; */
    }

    .grid {
        flex: 1;
        overflow-y: auto;
        padding: 10px;

        display: grid;
        grid-template-columns: repeat(auto-fill, minmax(180px, 1fr));
        align-content: start;
        gap: 10px;
    }

    .thumbnail {
        position: relative;
        aspect-ratio: 16 / 9;
        padding: 0;
        border: 2px solid transparent;
        border-radius: 4px;
        overflow: hidden;

        background-color: var(--primary-darkest);
        cursor: pointer;
    }
    .thumbnail:hover {
        filter: brightness(1.15);
    }
    .thumbnail.active {
        border-color: var(--secondary);
    }

    .thumbnail img {
        width: 100%;
        height: 100%;
        object-fit: contain;
    }

    .thumbnail .number {
        position: absolute;
        bottom: 4px;
        left: 4px;
        padding: 0 6px;
        border-radius: 3px;

        font-size: 0.8em;
        background-color: rgb(0 0 0 / 0.6);
        color: white;
    }

    .info {
        width: 100%;
        max-height: 60%;
        background-color: var(--primary-darkest);

        display: flex;
    }
</style>

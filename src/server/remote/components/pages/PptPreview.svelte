<script lang="ts">
    import Button from "../../../common/components/Button.svelte"
    import Icon from "../../../common/components/Icon.svelte"
    import { translate } from "../../util/helpers"
    import { send } from "../../util/socket"
    import { dictionary, mediaCache, outData } from "../../util/stores"

    export let active: any
    export let tablet: boolean = false

    $: path = active.id
    // state of the running slide show, sent with the output data
    $: presentation = $outData?.slide?.type === "ppt" && $outData.slide.id === path ? ($outData as any).presentation : null
    $: running = !!presentation

    $: slides = presentation?.stat?.slides || 0
    $: position = presentation?.stat?.position || 0
    $: thumbnails = (presentation?.thumbnails || []) as string[]
    $: notes = presentation?.info?.notes?.[position - 1] || ""

    // thumbnails are files on the computer, loaded once as images through the API
    const requested = new Set<string>()
    $: thumbnails.forEach((thumbnail) => {
        if (!thumbnail || $mediaCache[thumbnail] || requested.has(thumbnail)) return
        requested.add(thumbnail)
        send("API:get_thumbnail", { path: thumbnail })
    })

    function start() {
        send("API:start_presentation", { path })
    }

    function goto(index: number) {
        send("API:presentation_goto", { index })
    }
</script>

{#if !running}
    <div class="start">
        <Button on:click={start} variant="contained" center>
            <Icon id="play" size={1.5} right />
            {translate("presentation_control.start", $dictionary, "Start")}
        </Button>
    </div>
{:else}
    <div class="grid">
        {#each Array(slides) as _, i}
            {@const thumbnail = $mediaCache[thumbnails[i]]}
            <div class="main" class:active={position === i + 1} style="width: {100 / (tablet ? 3 : 2)}%;">
                <div class="slide" role="button" tabindex={0} on:click={() => goto(i + 1)} on:keydown={(e) => e.key === "Enter" && goto(i + 1)}>
                    {#if thumbnail}
                        <img src={thumbnail} alt="" />
                    {/if}
                    <span class="number">{i + 1}</span>
                </div>
            </div>
        {/each}
    </div>

    {#if notes}
        <p class="notes">{notes}</p>
    {/if}

    <div class="controls">
        <Button on:click={() => send("API:previous_slide")} variant="outlined" center compact>
            <Icon id="previous" size={1.2} />
        </Button>
        <span class="counter">
            {position}/{slides}
            {#if presentation.stat?.clickCount}
                <span style="opacity: 0.6;">({presentation.stat.clickIndex}/{presentation.stat.clickCount})</span>
            {/if}
        </span>
        <Button on:click={() => send("API:next_slide")} variant="outlined" center compact>
            <Icon id="next" size={1.2} />
        </Button>
    </div>
{/if}

<style>
    .start {
        flex: 1;
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 20px;
    }

    .grid {
        position: relative;
        flex: 1;
        overflow: auto;

        display: flex;
        flex-wrap: wrap;
        align-content: flex-start;
        padding: 5px;
        width: 100%;
    }

    .main {
        display: flex;
        position: relative;
        padding: 2px;
    }
    .main.active {
        outline: 2px solid var(--secondary);
        outline-offset: -1px;
    }

    .slide {
        position: relative;
        width: 100%;
        aspect-ratio: 16 / 9;
        overflow: hidden;

        display: flex;
        justify-content: center;
        align-items: center;

        background-color: var(--primary-darkest);
    }

    img {
        width: 100%;
        height: 100%;
        object-fit: contain;
        pointer-events: none;
    }

    .number {
        position: absolute;
        bottom: 4px;
        left: 4px;
        padding: 0 6px;
        border-radius: 3px;

        font-size: 0.8em;
        background-color: rgb(0 0 0 / 0.6);
        color: white;
    }

    .notes {
        max-height: 25%;
        overflow: auto;
        padding: 8px 12px;
        white-space: pre-wrap;
        background-color: var(--primary-darker);
    }

    .controls {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 10px;
        padding: 8px;
        background-color: var(--primary-darker);
    }

    .counter {
        font-weight: 600;
    }
</style>

<template>
    <footer>
        <button uxp-variant="cta" @click="onClickExecution">実行</button>
        <button type="submit" @click="onCancelHandler">キャンセル</button>
    </footer>
</template>
<script>
    export default {
        methods: {
            onCancelHandler() {
                this.$emit('onClickCancel');
            },
            onClickExecution() {
                this.$emit('onClickExecution');
            },
        },
    }
</script>
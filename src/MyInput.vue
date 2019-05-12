<template>
    <div>
    <p>{{ label }}</p>
    <input ref="textField" type="text" uxp-quiet="true">
    </div>
</template>
<script>
    export default {
        props: {
            label: {
                type: String,
                required: true,
            },
        },
        methods: {
            getValue: function(){
                return this.$refs.textField.value;
            },
        },
    }
</script>
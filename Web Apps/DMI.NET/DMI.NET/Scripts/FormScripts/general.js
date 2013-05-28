
function closeclick() {
    try {
        $(".popup").dialog("close");
        $("#optionframe").hide();
        $("#workframe").show();
    }
    catch (e) { }
}


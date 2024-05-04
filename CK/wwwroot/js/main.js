document.addEventListener('DOMContentLoaded', function () {
    var checkboxes = document.querySelectorAll('.multi-select-dropdown .form-check-input');
    var hiddenInput = document.getElementById('selectedStores');

    checkboxes.forEach(function (checkbox) {
        checkbox.addEventListener('change', function () {
            var selectedValues = Array.from(checkboxes)
                .filter(cb => cb.checked)
                .map(cb => cb.value);

            // Update the hidden input field with the selected values
            hiddenInput.value = selectedValues.join(',');
        });
    });
});
var matchEnterdDate = 0;
//function to set back date opacity for non supported browsers
window.onload = function () {
    var input = document.createElement('input');
    input.setAttribute('type', 'date');
    input.setAttribute('value', 'some text');
    if (input.value === "some text") {
        allDates = document.getElementsByClassName("xDateContainer");
        matchEnterdDate = 1;
        for (var i = 0; i < allDates.length; i++) {
            allDates[i].style.opacity = "1";
        }
    }
}
//function to convert enterd date to any format
document.addEventListener('DOMContentLoaded', function () {
    // Get references to the Quantity and Cost checkboxes
    var ItemLookupCodeCheckbox = document.querySelector('input[name="Parobj.VItemLookupCode"]');
    var ItemNameCheckbox = document.querySelector('input[name="Parobj.VItemName"]');
    var costCheckbox = document.querySelector('input[name="Parobj.VCost"]');
    var priceCheckbox = document.querySelector('input[name="Parobj.VPrice"]');

    function toggleCostCheckbox() {
        costCheckbox.disabled = !(ItemLookupCodeCheckbox.checked || ItemNameCheckbox.checked);
        priceCheckbox.disabled = !(ItemLookupCodeCheckbox.checked || ItemNameCheckbox.checked);
    }

    toggleCostCheckbox();

    ItemLookupCodeCheckbox.addEventListener('change', toggleCostCheckbox);
    ItemNameCheckbox.addEventListener('change', toggleCostCheckbox);
});
document.addEventListener('DOMContentLoaded', function () {
    var itemBarcodeCheckbox = document.getElementById('VItemLookupCode');
    var itemNameCheckbox = document.getElementById('VItemName');
    var costCheckbox = document.getElementById('VCost');
    var priceCheckbox = document.getElementById('VPrice');
    var displayLiItems = document.querySelectorAll('.Displayli');

    function isItemBarcodeCostOrPrice(element) {
        return (
            element === costCheckbox.parentElement ||
            element === priceCheckbox.parentElement ||
            element === itemBarcodeCheckbox.parentElement ||
            element === itemNameCheckbox.parentElement
        );
    }

    function updateHoverEffect() {
        var hoveredClass = 'hovered';

        displayLiItems.forEach(function (displayLi) {
            var label = displayLi.querySelector('label');

            if (!isItemBarcodeCostOrPrice(displayLi)) {
                label.classList.add(hoveredClass);
            } else {
                label.classList.remove(hoveredClass);
            }
        });
    }

    function handleLabelClick(event) {
        var checkbox = this.previousElementSibling;
        var displayLi = checkbox.parentElement;

        if (!isItemBarcodeCostOrPrice(displayLi)) {
            checkbox.checked = !checkbox.checked;
            event.preventDefault(); // Prevent the default action of the label click
        }
    }

    function enableLabelClickForItemName() {
        costCheckbox.parentElement.querySelector('label').classList.add('hovered');
        priceCheckbox.parentElement.querySelector('label').classList.add('hovered');
    }

    itemBarcodeCheckbox.addEventListener('change', function () {
        updateHoverEffect();
        if (itemBarcodeCheckbox.checked) {
            enableLabelClickForItemName();
        }
    });

    itemNameCheckbox.addEventListener('change', function () {
        updateHoverEffect();
        if (itemNameCheckbox.checked) {
            enableLabelClickForItemName();
        }
    });

    // displayLiItems.forEach(function (displayLi) {
    //     displayLi.addEventListener('mouseenter', function () {
    //         if (!isItemBarcodeCostOrPrice(displayLi)) {
    //             displayLi.classList.add('hovered');
    //         }
    //     });

    //     displayLi.addEventListener('mouseleave', function () {
    //         if (!isItemBarcodeCostOrPrice(displayLi)) {
    //             displayLi.classList.remove('hovered');
    //         }
    //     });

    //     // Add event listener for label clicks
    //     var label = displayLi.querySelector('label');
    //     label.addEventListener('click', handleLabelClick);
    // });

    // Initialize hover effect based on the initial state of the item barcode checkbox
    updateHoverEffect();
});
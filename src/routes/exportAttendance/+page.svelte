<script>
	import ComboBox from '../../lib/components/ComboBox.svelte';
	/** @type {import('./$types').PageData} */
	export let data;
	import { enhance } from '$app/forms';
	const { class_room_data } = data;
	const currentDate = new Date();
	// Từ ngày 10 của tháng trước ngày hiện tại
	const date1 = new Date(currentDate);
	date1.setDate(10); // Đặt ngày là ngày 10 của tháng hiện tại
	date1.setMonth(date1.getMonth() - 1); // Trừ một tháng để lấy tháng trước
	const date4 = new Date(currentDate);
	date4.setDate(9); // Đặt ngày là ngày 9 của tháng hiện tại
	date4.setMonth(date4.getMonth() + 1); // Cộng một tháng để lấy tháng sau

	function generateMonthData(numMonths) {
		const currentDate = new Date();
		let currentMonth = currentDate.getMonth() + 1; // Month is zero-based, so add 1
		let currentYear = currentDate.getFullYear();
		const monthsData = [];

		// Generate data for the current month
		monthsData.push({
			text: `${10}/${currentMonth}/${currentYear} - ${9}/${currentMonth + 1}/${currentYear}`,
			value: `${10}/${currentMonth}/${currentYear}`
		});

		// Generate data for previous months, maximum 3
		for (let i = 1; i <= 3; i++) {
			const previousMonth = currentMonth === 1 ? 12 : currentMonth - 1;
			const previousYear = currentMonth === 1 ? currentYear - 1 : currentYear;

			monthsData.unshift({
				text: `${10}/${previousMonth}/${previousYear} - ${9}/${currentMonth}/${currentYear}`,
				value: `${10}/${previousMonth}/${currentYear}`
			});

			// Update currentMonth and currentYear for the next iteration
			currentMonth = previousMonth;
			currentYear = previousYear;
		}
		currentMonth = currentDate.getMonth() + 2; // Month is zero-based, so add 1
		currentYear = currentDate.getFullYear();
		// Generate data for next months
		for (let i = 0; i < numMonths - 4; i++) {
			const nextMonth = currentMonth === 12 ? 1 : currentMonth + 1;
			const nextYear = currentMonth === 12 ? currentYear + 1 : currentYear;

			monthsData.push({
				text: `${10}/${currentMonth}/${currentYear} - ${9}/${nextMonth}/${nextYear}`,
				value: `${10}/${currentMonth}/${currentYear}`
			});

			// Update currentMonth and currentYear for the next iteration
			currentMonth = nextMonth;
			currentYear = nextYear;
		}

		return monthsData;
	}

	// Usage example:
	const fromDateData = generateMonthData(7);

	// Helper function to convert ExcelJS workbook to base64
	export let form;
</script>

<h1>Tải file</h1>

<div class="form">
	{#if form?.error == true}
		<h2>{form?.message}</h2>
	{/if}

	{#if form?.error === false}
		<h2>{form?.message}</h2>
	{/if}
	<form
		method="post"
		action="?/download"
		use:enhance={({ formElement, formData, action, cancel, submitter }) => {
			// `formElement` is this `<form>` element
			// `formData` is its `FormData` object that's about to be submitted
			// `action` is the URL to which the form is posted
			// calling `cancel()` will prevent the submission
			// `submitter` is the `HTMLElement` that caused the form to be submitted

			return async ({ result, update }) => {
				let buffer = result.data['workbook'];
				let filename = result.data['filename'];
				if (buffer == null || filename == null) {
					update({ reset: false });
				} else {
					// Create a blob from the base64 string
					const byteCharacters = atob(buffer);
					const byteNumbers = new Array(byteCharacters.length);
					for (let i = 0; i < byteCharacters.length; i++) {
						byteNumbers[i] = byteCharacters.charCodeAt(i);
					}
					const byteArray = new Uint8Array(byteNumbers);
					const blob = new Blob([byteArray], {
						type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,'
					});

					// Create a temporary URL for the blob
					const url = window.URL.createObjectURL(blob);

					// Create a link element to trigger the download
					const a = document.createElement('a');
					a.href = url;
					a.download = filename;
					a.style.display = 'none';

					// Append the link to the DOM and trigger the click event
					document.body.appendChild(a);
					a.click();

					// Clean up by removing the link and revoking the blob URL
					document.body.removeChild(a);
					window.URL.revokeObjectURL(url);
					update({ reset: false });
				}
			};
		}}
	>
		<div class="stack">
			<ComboBox
				label="Chọn lớp học"
				name="class_room_id"
				placeholder="Type to search..."
				options={class_room_data}
			/>

			<ComboBox
				label="Chọn khoảng thời gian"
				name="fromDate"
				placeholder="Type to search..."
				options={fromDateData}
				readonly={false}
			/>
			<button type="submit">Tải file điểm danh</button>
		</div>
	</form>
</div>

<style>
	.stack {
		display: flex;
		flex-direction: column;
		gap: 1.5rem;
	}
</style>

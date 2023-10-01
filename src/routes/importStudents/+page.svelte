<script>
	/** @type {import('./$types').PageData} */
	export let data;
	const { class_room_data } = data;
	import ComboBox from '../../lib/components/ComboBox.svelte';

	export let form;
</script>

{#if form?.error}
	<h2>{form?.message}</h2>
{/if}

{#if form?.success}
	<h2>File Uploaded!</h2>
{/if}
<h1>Upload File</h1>
<div class="form">
	<form method="post" action="?/upload" enctype="multipart/form-data">
		<label for="file-upload">Chọn file:</label>
		<input type="file" id="file-upload" name="file-upload" accept="*/*" required />
		<label for="worksheet-name">Chọn tên của worksheet chứa danh sách học sinh:</label>
		<input type="text" id="worksheet-name" name="worksheet_name" placeholder="Sheet1" required />
		<div class="stack">
			<ComboBox
				label="Chọn lớp học"
				name="class_room_id"
				placeholder="Type to search..."
				options={class_room_data}
			/>
		</div>

		<button type="submit">Upload</button>
	</form>
</div>

<style>
	.stack {
		display: flex;
		flex-direction: column;
		gap: 1.5rem;
	}
</style>

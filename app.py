st.subheader("Saved Entries (Edit & Delete Directly)")
    st.caption("📝 **To Edit:** Double-click any cell to change its value. <br> 🗑️ **To Delete:** Click the grey box on the far left of a row to select it, then press the **Delete / Trash** icon in the top right of the table.", unsafe_allow_html=True)

    if filtered.empty:
        st.info("No entries saved for this employee and period.")
    else:
        # Display the interactive data editor
        edited_df = st.data_editor(
            filtered,
            use_container_width=True,
            hide_index=False,
            num_rows="dynamic", # This allows row deletion
            disabled=["id", "total", "created_at", "empName", "designation", "location", "reportMonthText"], # Lock background columns
            key="expense_editor"
        )

        # Button to save the direct table edits to Google Sheets
        if st.button("💾 Save Table Changes to Google Sheets", type="primary"):
            try:
                # Recalculate totals just in case the user edited the money columns
                for idx, row in edited_df.iterrows():
                    edited_df.at[idx, "total"] = calc_total(row)

                # Fetch the main database
                full_df = read_entries()

                # Remove the old versions of this employee's current month records
                mask = (full_df["empName"] == employee["empName"]) & (full_df["reportMonthText"] == employee["reportMonthText"])
                full_df = full_df[~mask]

                # Add the newly edited/deleted records back in
                final_df = pd.concat([full_df, edited_df], ignore_index=True)

                # Overwrite Google Sheets with the updated data
                rewrite_entries(final_df)

                st.success("All edits and deletions have been successfully synced to Google Sheets!")
                st.rerun()
            except Exception as e:
                st.error(f"Failed to save changes: {e}")
                with st.expander("Show full error"):
                    st.code(traceback.format_exc())

if __name__ == "__main__":
    main()

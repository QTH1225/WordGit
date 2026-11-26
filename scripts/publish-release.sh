#!/usr/bin/env bash
set -euo pipefail

assets=(release/*)
[[ ${#assets[@]} -eq 6 ]]
release_id() {
    gh api --paginate "repos/$GH_REPO/releases?per_page=100" \
        --jq ".[] | select(.tag_name == \"$RELEASE_TAG\") | .id"
}

id=$(release_id)
if [[ -z $id ]]; then
    gh release create "$RELEASE_TAG" --verify-tag --draft --title "WordGit $VERSION" \
        --notes-file release-notes.md --prerelease="$PRERELEASE" --latest=false
    id=$(release_id)
fi
[[ -n $id ]]
endpoint="repos/$GH_REPO/releases/$id"
state=$(gh api "$endpoint")
jq -e --arg tag "$RELEASE_TAG" --arg flag "$PRERELEASE" \
    '.tag_name == $tag and .draft and (.prerelease | tostring) == $flag' <<< "$state" >/dev/null

for file in "${assets[@]}"; do
    name=${file##*/}
    digest=$(sha256sum "$file")
    digest="sha256:${digest%% *}"
    asset=$(jq -c --arg name "$name" '.assets[] | select(.name == $name)' <<< "$state")
    if [[ -z $asset ]]; then
        gh release upload "$RELEASE_TAG" "$file"
    elif ! jq -e --arg digest "$digest" '.state == "uploaded" and .digest == $digest' <<< "$asset" >/dev/null; then
        printf 'Draft asset differs from this build: %s\n' "$name" >&2
        exit 1
    fi
done

state=$(gh api "$endpoint")
jq -e '.assets | length == 6' <<< "$state" >/dev/null
for file in "${assets[@]}"; do
    digest=$(sha256sum "$file")
    digest="sha256:${digest%% *}"
    jq -e --arg name "${file##*/}" --arg digest "$digest" \
        'any(.assets[]; .name == $name and .state == "uploaded" and .digest == $digest)' <<< "$state" >/dev/null
done
if [[ $PRERELEASE == true ]]; then
    gh release edit "$RELEASE_TAG" --draft=false --prerelease --latest=false
else
    gh release edit "$RELEASE_TAG" --draft=false --prerelease=false --latest
fi

$ids = @('EPPlus','EPPlus.Interfaces','Microsoft.Bcl.Cryptography','Microsoft.IO.RecyclableMemoryStream','System.Buffers','System.ComponentModel.Annotations','System.Formats.Asn1','System.Memory','System.Numerics.Vectors','System.Runtime.CompilerServices.Unsafe','System.Security.Cryptography.Xml','System.ValueTuple')
foreach ($id in $ids) {
	$url = "https://api.nuget.org/v3/registration5-semver1/$($id.ToLower())/index.json"
	try {
		$json = Invoke-RestMethod -Uri $url -ErrorAction Stop
		$versions = @()
		foreach ($page in $json.items) {
			foreach ($item in $page.items) {
				$v = $item.catalogEntry.version
				if ($v -notmatch '-') { $versions += [version]$v }
			}
		}
		if ($versions.Count -gt 0) {
			$max = ($versions | Sort-Object)[-1].ToString()
			Write-Output "$id|$max"
		} else {
			Write-Output "$id|NO_STABLE"
		}
	} catch {
		Write-Output "$id|ERROR:$($_.Exception.Message)"
	}
}

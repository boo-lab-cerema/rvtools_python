from pyVmomi import vim

# from rvtools.printrv.csv_print import *
from rvtools.printrv.xls_print import xls_print


def get_obj(content, vimtype):
    obj = None
    container = content.viewManager.CreateContainerView(
        content.rootFolder, vimtype, True
    )
    obj = container.view[0].name
    # for c in container.view:
    #     if name:
    #         if c.name == name:
    #             obj = c
    #             break
    #     else:
    #         obj = c
    #         break
    return obj


def vinfo_collect(service_instance, directory):
    """Def responsible to connect on the vCenter and retrieve VM information"""

    print("## Processing vInfo module")

    content = service_instance.RetrieveContent()
    container = content.rootFolder

    object_list = []

    #############################################
    # Datastore objects                         #
    #############################################
    print("=====================")
    print("Datastores:")
    print("=====================")
    view_type = [vim.Datastore]
    container_view = content.viewManager.CreateContainerView(container, view_type, True)
    children = container_view.view
    for child in children:
        summary = child.summary
        vinfo_data = []
        vinfo = {"object": summary.name or "", "data": vinfo_data}
        vinfo_data.append([{"capacity (bytes)": summary.capacity}])
        vinfo_data.append([{"free space (bytes)": summary.freeSpace}])
        vinfo_data.append(
            [{"used (computed, bytes)": summary.capacity - summary.freeSpace}]
        )
        vinfo_data.append([{"type": summary.type}])

        # TODO
        # ds.info : vim.Datastore.Info

        print("Datastore: {}".format(summary.name))
        object_list.append(vinfo)

    #############################################
    # VitrualMachine objects                    #
    #############################################
    print("=====================")
    print("VMs:")
    print("=====================")
    view_type = [vim.VirtualMachine]
    container_view = content.viewManager.CreateContainerView(container, view_type, True)

    children = container_view.view
    for child in children:
        # if 'sat62' in child.name:
        # if 'akavir-sat63' in child.name:
        # if 'waldirio' in child.name:
        if True:
            vinfo_data = []
            vinfo = {"object": child.name or "", "data": vinfo_data}

            # storage.perDatastoreUsage
            perDatastoreUsage = child.storage.perDatastoreUsage
            dsinfo_data = []
            for usage in perDatastoreUsage:
                per_ds_data = {}
                datastoreInfos = usage.datastore.info
                datastoreName = datastoreInfos.name
                per_ds_data["Datastore"] = datastoreName
                # Storage space, in bytes, on this datastore that is actually being used by the virtual machine.
                per_ds_data["DS space used by vm (bytes)"] = usage.committed
                #  Additional storage space, in bytes, potentially used by the virtual machine on this datastore. (e.g. lazily allocated disks grow, or storage for swap...) -- /!\  If the virtual machine is running off delta disks (for example because a snapshot was taken), then only the potential growth of the currently used delta-disks is considered.
                per_ds_data["DS potential growth by vm (bytes)"] = usage.uncommitted
                #  Storage space, in bytes, occupied by the virtual machine on this datastore that is not shared with any other virtual machine.
                per_ds_data["DS occupied and not shared by vm (bytes)"] = usage.unshared
                dsinfo_data.append(per_ds_data)
            vinfo_data.append(dsinfo_data)

            # runtime.powerState
            powerstate = child.runtime.powerState
            if powerstate is None:
                powerstate = ""
            vinfo_data.append([{"powerstate": powerstate}])

            # configStatus
            config_status = child.configStatus
            if config_status is None:
                config_status = ""
            vinfo_data.append([{"config_status": config_status}])

            # guest.hostName
            dns_name = child.guest.hostName
            if dns_name is None:
                dns_name = ""
            vinfo_data.append([{"dns_name": dns_name}])

            # runtime.connectionState
            connection_state = child.runtime.connectionState
            if connection_state is None:
                connection_state = ""
            vinfo_data.append([{"connection_state": connection_state}])

            # guest.guestState
            guest_state = child.guest.guestState
            if guest_state is None:
                guest_state = ""
            vinfo_data.append([{"guest_state": guest_state}])

            # config.hardware.numCPU
            cpus = child.config.hardware.numCPU
            if cpus is None:
                cpus = ""
            vinfo_data.append([{"cpus": cpus}])

            # config.hardware.memoryMB
            memory = child.config.hardware.memoryMB
            if memory is None:
                memory = ""
            vinfo_data.append([{"memory": memory}])

            # network count
            nics = child.network.__len__()
            vinfo_data.append([{"nics": nics}])

            # layout.disk count
            disks = child.layout.disk.__len__()
            if disks is None:
                disks = ""
            vinfo_data.append([{"disks": disks}])

            # config.firmware
            firmware = child.config.firmware
            if firmware is None:
                firmware = ""
            vinfo_data.append([{"firmware": firmware}])

            # config.files.vmPathName
            path = child.config.files.vmPathName
            if path is None:
                path = ""
            vinfo_data.append([{"path": path}])

            # Datacenter
            datacenter = get_obj(content, [vim.Datacenter])
            vinfo_data.append([{"datacenter": datacenter}])

            # ClusterComputeResource
            cluster = get_obj(content, [vim.ClusterComputeResource])
            vinfo_data.append([{"cluster": cluster}])

            # config.guestFullName
            os_according_to_the_vmware_tools = child.config.guestFullName
            if os_according_to_the_vmware_tools is None:
                os_according_to_the_vmware_tools = ""
            vinfo_data.append(
                [{"os_according_to_the_vmware_tools": os_according_to_the_vmware_tools}]
            )

            # vm id
            vm_id = child._moId
            if vm_id is None:
                vm_id = ""
            vinfo_data.append([{"vm_id": vm_id}])

            # config.uuid
            vm_uuid = child.config.uuid
            if vm_uuid is None:
                vm_uuid = ""
            vinfo_data.append([{"vm_uuid": vm_uuid}])

            print("VM: {}".format(child.name))

            object_list.append(vinfo)

    xls_print("vinfo.xlsx", object_list, directory)
